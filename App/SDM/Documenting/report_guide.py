"""
This module contains a class for managing column descriptions used in exported workbooks.
"""

import pandas as pd
from collections import OrderedDict

class ReportGuide:
    def __init__(self):
        # Base curve feature descriptions (used in Features tab)
        self._region_definitions = {
            '_CURVE': 'The main TAC curve region where the alcohol concentration rises and falls',
            '_PERIPHERY': 'The region before and after the curve, used for quality assessment',
            '_PERIPHERY_BEFORE': 'The region before the curve start, used for quality assessment',
            '_PERIPHERY_AFTER': 'The region after the curve end, used for quality assessment',
            '_REGION': 'The region consisting of CURVE + PERIPHERY',
            '_EMA_REGION': 'The 6-hour region around the earliest timestamp of the event (2 hours before, 4 hours after)',
        }
        self._curve_feature_descriptions = {
            # Basic identifiers
            'subid': 'Subject identifier',
            'dataset_id': 'Dataset identifier',
            'curve_id': 'Curve identifier',
            'curve_id_imputed_match': 'Imputed curve_id with highest temporal overlap to this raw curve (independent raw demarcation)',
            'imputed_match_overlap_percent': 'Fraction of raw curve window overlapped by the matched imputed curve (0-1)',
            'threshold_source': 'Whether threshold metadata on this row came from raw or imputed k-means clustering',
            'curve_count': 'Number of curves for this subject',
            'curve_threshold': 'Threshold used for curve detection',
            
            # Device information
            'device_one_REGION': 'First device ID in the region',
            'device_two_REGION': 'Second device ID in the region',
            'device_count': 'Number of devices in the region',
            
            # TAC curve features
            'duration': 'Total duration in hours',
            'auc_total': 'Area under the curve (total) (µg/L·min)',
            'auc_relative': 'Area under the curve (relative to baseline) (µg/L·min)',
            'peak': 'Maximum TAC value (µg/L)',
            'relative_peak': 'Maximum TAC value relative to baseline (µg/L)',
            'rise_fall_rate_CURVE': 'Ratio of curve duration to relative peak (hours per µg/L) - indicates how much overall rise-to-fall movement occurs across time',
            'rise_rate_CURVE': 'Rate of TAC increase from curve threshold to peak (µg/(L·h))',
            'rise_rate_point_to_point_CURVE': 'Average rate of ascending point-to-point TAC changes (µg/(L·min))',
            'rise_duration_point_to_point_CURVE': 'Total duration of ascending point-to-point TAC changes (hours)',
            'rise_rate_1hr_CURVE': 'Rate of TAC increase over first hour, bounded by peak (µg/(L·h))',
            'rise_rate_2hr_CURVE': 'Rate of TAC increase over first 2 hours, bounded by peak (µg/(L·h))',
            'fall_rate_CURVE': 'Rate of TAC decrease from peak to curve end (µg/(L·h))',
            'fall_rate_point_to_point_CURVE': 'Average rate of descending point-to-point TAC changes (µg/(L·min))',
            'fall_duration_point_to_point_CURVE': 'Total duration of descending point-to-point TAC changes (hours)',
            'fall_rate_1hr_CURVE': 'Rate of TAC decrease over last hour, bounded by peak (µg/(L·h))',
            'fall_rate_2hr_CURVE': 'Rate of TAC decrease over last 2 hours, bounded by peak (µg/(L·h))',
            'rise_duration': 'Duration of TAC increase (hours)',
            'fall_duration': 'Duration of TAC decrease (hours)',
            'ascending_imputed_percent_CURVE': 'Fraction of ascending TAC pairs containing imputed values',
            'descending_imputed_percent_CURVE': 'Fraction of descending TAC pairs containing imputed values',
            'CURVE_VALID': 'Whether the curve meets validity criteria (1=valid)',
            'REGION_VALID': 'Whether both the curve and periphery meet validity criteria (1=valid)',
            'DRINKING_PRED': 'Predicted drinking curve: duration_CURVE > default 29 min, high_quality_duration_CURVE > required_HQ_duration (four-phase vs curve length), FLAG_below_threshold_curve==0, FLAG_incomplete_curve_start_curve==0 (1=likely drinking)',
            
            # Curve timing features
            'begin': 'Start time of the curve',
            'end': 'End time of the curve',
            'first_tac': 'First TAC value in the curve',
            'last_tac': 'Last TAC value in the curve',
            'mean_tac': 'Mean TAC value in the curve',
            'sd_tac': 'Standard deviation of TAC values in the curve',
            'sem_tac': 'Standard error of the mean TAC values in the curve',
            'rise_complete_perc': 'Percentage of rise phase that is complete (peak - first_tac)/relative_peak',
            'fall_complete_perc': 'Percentage of fall phase that is complete (peak - last_tac)/relative_peak',
            
            # Quality features
            'total_duration': 'Total duration of the region (hours)',
            'device_turned_on_duration': 'Duration device was turned on (hours)',
            'device_turned_on_percent': 'Percentage of time device was turned on',
            'device_worn_duration': 'Duration device was worn (hours)',
            'device_worn_percent': 'Percentage of time device was worn',
            'flatline_max': 'Maximum duration of flatline (hours)',
            'flatlined_percent': 'Percentage of flatlined data',
            'high_quality_duration_CURVE': 'Duration of high-quality data in the curve (hours)',
            'high_quality_percent_CURVE': 'Percentage of high-quality data in the curve',
            'required_HQ_duration': 'Required high-quality duration for drinking prediction (hours): four-phase algorithm - Phase 1 (≤30 min): 100% HQ required; Phase 2 (30-60 min): 6 additional HQ minutes per 15-min block; Phase 3 (60-150 min): 7 additional HQ minutes per 15-min block; Phase 4 (150+ min): percentage-based starting at 56%, declining 2% per 30 min, floor at 40%',
            'high_quality_above_threshold_duration_CURVE': 'Duration of high-quality TAC values at or above the curve threshold (hours)',
            
            # Low quality features
            'imputed_low_quality_duration': 'Duration of imputed low quality data (hours)',
            'imputed_low_quality_percent': 'Percentage of imputed low quality data',
            'unimputed_low_quality_duration': 'Duration of low quality data not imputed (hours)',
            'unimputed_low_quality_percent': 'Percentage of low quality data not imputed',
            'total_low_quality_duration': 'Total duration of low quality data (hours)',
            'total_low_quality_percent': 'Total percentage of low quality data',
            
            # Gap features
            'imputed_gap_duration': 'Duration of imputed gap data (hours)',
            'imputed_gap_percent': 'Percentage of imputed gap data',
            'unimputed_gap_duration': 'Duration of unimputed gap data (hours)',
            'unimputed_gap_percent': 'Percentage of unimputed gap data',
            'total_gap_duration': 'Total duration of gap data (hours)',
            'total_gap_percent': 'Total percentage of gap data',
            'total_gap_duration_inclusive': 'Total duration of gap data (hours); inclusive: gap==1 only, no mutual exclusivity',
            'total_gap_percent_inclusive': 'Total percentage of gap data; inclusive: gap==1 only, no mutual exclusivity',
            'gap_imputation_ratio': 'Ratio of gap data that has been imputed (0-1)',
            
            # Non-wear features
            'imputed_non_wear_duration': 'Duration of imputed non-wear data (hours)',
            'imputed_non_wear_percent': 'Percentage of imputed non-wear data',
            'unimputed_non_wear_duration': 'Duration of unimputed non-wear data (hours)',
            'unimputed_non_wear_percent': 'Percentage of unimputed non-wear data',
            'total_non_wear_duration': 'Total duration of non-wear data (hours)',
            'total_non_wear_percent': 'Total percentage of non-wear data',
            'total_non_wear_duration_inclusive': 'Total duration of non-wear data (hours); inclusive: non_wear==1 only, no mutual exclusivity',
            'total_non_wear_percent_inclusive': 'Total percentage of non-wear data; inclusive: non_wear==1 only, no mutual exclusivity',
            'non_wear_imputation_ratio': 'Ratio of non-wear data that has been imputed (0-1)',
            
            # Jump features
            'imputed_jump_duration_CURVE': 'Duration of imputed jump data (hours)',
            'imputed_jump_percent_CURVE': 'Percentage of imputed jump data',
            'unimputed_jump_duration_CURVE': 'Duration of unimputed jump data (hours)',
            'unimputed_jump_percent_CURVE': 'Percentage of unimputed jump data',
            'total_jump_duration_CURVE': 'Total duration of jump data (hours)',
            'total_jump_percent_CURVE': 'Total percentage of jump data',
            'total_jump_duration_inclusive': 'Total duration of jump data (hours); inclusive: jump==1 only, no mutual exclusivity',
            'total_jump_percent_inclusive': 'Total percentage of jump data; inclusive: jump==1 only, no mutual exclusivity',
            'jump_imputation_ratio_CURVE': 'Ratio of jump data that has been imputed (0-1)',
            
            # Plummet features
            'imputed_plummet_duration': 'Duration of imputed plummet data (hours)',
            'imputed_plummet_percent': 'Percentage of imputed plummet data',
            'unimputed_plummet_duration': 'Duration of unimputed plummet data (hours)',
            'unimputed_plummet_percent': 'Percentage of unimputed plummet data',
            'total_plummet_duration': 'Total duration of plummet data (hours)',
            'total_plummet_percent': 'Total percentage of plummet data',
            'total_plummet_duration_inclusive': 'Total duration of plummet data (hours); inclusive: plummet==1 only, no mutual exclusivity',
            'total_plummet_percent_inclusive': 'Total percentage of plummet data; inclusive: plummet==1 only, no mutual exclusivity',
            'plummet_imputation_ratio': 'Ratio of plummet data that has been imputed (0-1)',
            
            # Extreme negative features
            'imputed_extreme_negative_duration': 'Duration of imputed extreme negative data (hours)',
            'imputed_extreme_negative_percent': 'Percentage of imputed extreme negative data',
            'unimputed_extreme_negative_duration': 'Duration of unimputed extreme negative data (hours)',
            'unimputed_extreme_negative_percent': 'Percentage of unimputed extreme negative data',
            'total_extreme_negative_duration': 'Total duration of extreme negative data (hours)',
            'total_extreme_negative_percent': 'Total percentage of extreme negative data',
            'total_extreme_negative_duration_inclusive': 'Total duration of extreme negative data (hours); inclusive: extreme_negative==1 only, no mutual exclusivity',
            'total_extreme_negative_percent_inclusive': 'Total percentage of extreme negative data; inclusive: extreme_negative==1 only, no mutual exclusivity',
            'sub_negative_10_sum': 'Sum of all TAC values below -10',
            'extreme_negative_imputation_ratio': 'Ratio of extreme negative data that has been imputed (0-1)',
            'low_quality_imputation_ratio': 'Ratio of overall low quality data that has been imputed (0-1)',
            
            # Curve counts
            'started_curve_count': 'Number of discrete curves within the final curve',
            'complete_curve_count': 'Number of discrete curves started and ended within the final curve',
            
            # Visualization paths
            'smoothed_curve_plot': 'Path to smoothed curve plot',
            'signal_processing_plot': 'Path to signal processing plot',
            'device_removal_plot': 'Path to device removal plot',
            'signal_processing_plot_wide': 'Path to wide signal processing plot',
            
            # Validity flags
            'FLAG_sub_negative_X_PERIPHERY_>P%_>Yhrs': 'Flag for >P% of values below -X for >Y hours in periphery',
            'FLAG_sub_negative_X_PERIPHERY_>Q%_>Yhrs': 'Flag for >Q% of values below -X for >Y hours in periphery',
            'FLAG_sub_negative_X_PERIPHERY_>R%_>Yhrs': 'Flag for >R% of values below -X for >Y hours in periphery',
            'FLAG_non_wear_PERIPHERY_>P%': 'Flag for >P% non-wear in periphery',
            'FLAG_flatlined_peak_CURVE_>P%flatline_peak>X': 'Flag for >P% flatline with peak >X',
            'FLAG_sub_negative_X_CURVE_>P%_>Y': 'Flag for >P% of values below -X for >Y hours in curve',
            'FLAG_rise_completion_CURVE_<P%': 'Flag for <P% rise completion',
            'FLAG_rise_rate_CURVE_>X': 'Flag for threshold-to-peak rise rate >X',
            'FLAG_fall_completion_CURVE_<P%': 'Flag for <P% fall completion',
            'FLAG_short_curve_duration_CURVE_<Xhrs': 'Flag for curve duration <X hours',
            'FLAG_imputed_CURVE_>P%_or_duration>Xhrs': 'Flag for >P% imputed or duration >X hours',
            'FLAG_unimputed_low_quality_CURVE_>P%': 'Flag for >P% unimputed low quality data',
            
            # Periphery flags
            'FLAG_gaps_and_non_wear_periphery_before': 'Flag for gaps and non-wear in periphery before exceeding threshold',
            'FLAG_extreme_negative_periphery_before': 'Flag for extreme negative values in periphery before exceeding threshold',
            'FLAG_low_quality_periphery_before': 'Flag for low quality data in periphery before exceeding threshold',
            'FLAG_unimputed_low_quality_periphery_before': 'Flag for unimputed low quality data in periphery before exceeding threshold',
            'FLAG_gaps_and_non_wear_periphery_after': 'Flag for gaps and non-wear in periphery after exceeding threshold',
            'FLAG_extreme_negative_periphery_after': 'Flag for extreme negative values in periphery after exceeding threshold',
            'FLAG_low_quality_periphery_after': 'Flag for low quality data in periphery after exceeding threshold',
            'FLAG_unimputed_low_quality_periphery_after': 'Flag for unimputed low quality data in periphery after exceeding threshold',
            
            # Validity indicators
            'PERIPHERY_VALID': 'Whether the periphery meets validity criteria (1=valid)',
            
            # Baseline information
            'unadjusted_threshold': 'Unadjusted threshold value',
            'baseline_mean': 'Mean baseline value',
            'baseline_sd': 'Standard deviation of baseline values',
            
            # Event matching information
            'CURVE_event_match_before_buffer': 'Buffer time before curve for event matching',
            'CURVE_event_match_after_buffer': 'Buffer time after curve for event matching',
            'CURVE_MATCH_START': 'Start time of curve match window',
            'CURVE_MATCH_END': 'End time of curve match window',
            'event_matched_1': 'First matched event ID',
            'event_matched_2': 'Second matched event ID',
            'event_matched_3': 'Third matched event ID',
            'event_matched_4': 'Fourth matched event ID',
            'event_matched_5': 'Fifth matched event ID',
            'event_matched_6': 'Sixth matched event ID',
            'event_matched_7': 'Seventh matched event ID',
            'event_matched_8': 'Eighth matched event ID',
        }

        # Event-related feature descriptions (used in Events tab)
        self._event_feature_descriptions = {
            # Basic event identifiers
            'subid': 'Subject identifier',
            'dataset_identifier': 'Dataset identifier',
            'ema_id': 'EMA (Ecological Momentary Assessment) identifier',
            'event': 'Event number',
            'day_id': 'Study day identifier',
            'drink_total': 'Total number of drinks reported',
            
            # Event timestamps
            'earliest_timestamp': 'Earliest timestamp associated with the event',
            'latest_timestamp': 'Latest timestamp associated with the event',
            'matching_end_timestamp': 'End timestamp for event matching window',
            'self_report_start_time': 'Start time of self-reported event',
            
            # Curve matching information
            'matched': 'Whether the event matched to any curve (1=matched)',
            'num_curves_matched': 'Number of curves matched to this event',
            'has_shared_match': 'Whether the event shares a valid curve match with another event',
            'shared_curve_id': 'ID of the shared valid curve if event shares a match',
            'has_shared_first_match': 'Whether the event shares its first valid curve match with another event',
            'shared_first_curve_id': 'ID of the first shared valid curve if event shares a first match',
            
            # Curve match details (up to 5 possible matches)
            'curve_match_1': 'ID of the first matched curve',
            'curve_match_1_overlap': 'Proportion of overlap between event and first matched curve',
            'curve_match_2': 'ID of the second matched curve',
            'curve_match_2_overlap': 'Proportion of overlap between event and second matched curve',
            'curve_match_3': 'ID of the third matched curve',
            'curve_match_3_overlap': 'Proportion of overlap between event and third matched curve',
            'curve_match_4': 'ID of the fourth matched curve',
            'curve_match_4_overlap': 'Proportion of overlap between event and fourth matched curve',
            'curve_match_5': 'ID of the fifth matched curve',
            'curve_match_5_overlap': 'Proportion of overlap between event and fifth matched curve',
            
            # Valid curve match details
            'valid_curve_match_1': 'ID of the first valid matched curve',
            'valid_curve_match_2': 'ID of the second valid matched curve',
            'valid_curve_match_3': 'ID of the third valid matched curve',
            'valid_curve_match_4': 'ID of the fourth valid matched curve',
            'valid_curve_match_5': 'ID of the fifth valid matched curve',
            'valid_curve_match_1_overlap': 'ID of the first valid curve that overlaps with another valid curve',
            'valid_curve_match_2_overlap': 'ID of the second valid curve that overlaps with another valid curve',
            'valid_curve_match_3_overlap': 'ID of the third valid curve that overlaps with another valid curve',
            'valid_curve_match_4_overlap': 'ID of the fourth valid curve that overlaps with another valid curve',
            'valid_curve_match_5_overlap': 'ID of the fifth valid curve that overlaps with another valid curve',
            
            # Invalid curve match details
            'invalid_curve_match_1': 'ID of the first invalid matched curve',
            'invalid_curve_match_2': 'ID of the second invalid matched curve',
            'invalid_curve_match_3': 'ID of the third invalid matched curve',
            'invalid_curve_match_4': 'ID of the fourth invalid matched curve',
            'invalid_curve_match_5': 'ID of the fifth invalid matched curve',
            'invalid_curve_match_1_overlap': 'ID of the first invalid curve that overlaps with another invalid curve',
            'invalid_curve_match_2_overlap': 'ID of the second invalid curve that overlaps with another invalid curve',
            'invalid_curve_match_3_overlap': 'ID of the third invalid curve that overlaps with another invalid curve',
            'invalid_curve_match_4_overlap': 'ID of the fourth invalid curve that overlaps with another invalid curve',
            'invalid_curve_match_5_overlap': 'ID of the fifth invalid curve that overlaps with another invalid curve',
            
            # Curve validity information
            'CURVE_and_PERIPHERY_VALID_1': 'Whether the first matched curve and its periphery are valid (1=valid)',
            'CURVE_and_PERIPHERY_VALID_2': 'Whether the second matched curve and its periphery are valid (1=valid)',
            'CURVE_and_PERIPHERY_VALID_3': 'Whether the third matched curve and its periphery are valid (1=valid)',
            'CURVE_and_PERIPHERY_VALID_4': 'Whether the fourth matched curve and its periphery are valid (1=valid)',
            'CURVE_and_PERIPHERY_VALID_5': 'Whether the fifth matched curve and its periphery are valid (1=valid)',
            
            # Event region quality features
            'total_duration_EMA_REGION': 'Total duration of the EMA region (hours)',
            'TAC_max_EMA_REGION': 'Maximum TAC value in the EMA region (µg/L)',
            'TAC_avg_EMA_REGION': 'Average TAC value in the EMA region (µg/L)',
            'device_turned_on_duration_EMA_REGION': 'Duration device was turned on in EMA region (hours)',
            'device_turned_on_percent_EMA_REGION': 'Percentage of time device was turned on in EMA region',
            'device_worn_duration_EMA_REGION': 'Duration device was worn in EMA region (hours)',
            'device_worn_percent_EMA_REGION': 'Percentage of time device was worn in EMA region',
            'imputed_duration_EMA_REGION': 'Duration of imputed data in EMA region (hours)',
            'imputed_percent_EMA_REGION': 'Percentage of data that was imputed in EMA region',
            'low_quality_duration_EMA_REGION': 'Duration of low quality data in EMA region (hours)',
            'low_quality_percent_EMA_REGION': 'Percentage of low quality data in EMA region',
            'unimputed_low_quality_duration_EMA_REGION': 'Duration of low quality data not imputed in EMA region (hours)',
            'unimputed_low_quality_percent_EMA_REGION': 'Percentage of low quality data not imputed in EMA region',
            
            # Event region negative value features
            'negative_duration_EMA_REGION': 'Duration of negative TAC values in EMA region (hours)',
            'sub_negative_10_duration_EMA_REGION': 'Duration of TAC values below -10 in EMA region (hours)',
            'sub_negative_10_percent_EMA_REGION': 'Percentage of TAC values below -10 in EMA region',
            'consecutive_sub_negative_10_duration_EMA_REGION': 'Longest consecutive duration of TAC values below -10 in EMA region (hours)',
            'sub_negative_20_duration_EMA_REGION': 'Duration of TAC values below -20 in EMA region (hours)',
            'sub_negative_20_percent_EMA_REGION': 'Percentage of TAC values below -20 in EMA region',
            'consecutive_sub_negative_20_duration_EMA_REGION': 'Longest consecutive duration of TAC values below -20 in EMA region (hours)',
            'sub_negative_40_duration_EMA_REGION': 'Duration of TAC values below -40 in EMA region (hours)',
            'sub_negative_40_percent_EMA_REGION': 'Percentage of TAC values below -40 in EMA region',
            'consecutive_sub_negative_40_duration_EMA_REGION': 'Longest consecutive duration of TAC values below -40 in EMA region (hours)',
            
            # Event region visualization
            'device_removal_plot_EMA_REGION': 'Path to device removal plot for EMA region',
            'signal_processing_plot_EMA_REGION': 'Path to signal processing plot for EMA region',
            
        }

        # Person-level feature descriptions (used in Valid Matched by SubID tab)
        self._person_level_feature_descriptions = {
            # Continuous TAC features
            'duration_CURVE': 'Total duration in hours',
            'auc_total_CURVE': 'Area under the curve (total) (µg/L·min)',
            'auc_relative_CURVE': 'Area under the curve (relative to baseline) (µg/L·min)',
            'peak_CURVE': 'Maximum TAC value (µg/L)',
            'relative_peak_CURVE': 'Maximum TAC value relative to baseline (µg/L)',
            'rise_fall_rate_CURVE': 'Ratio of curve duration to relative peak (hours per µg/L) - indicates how much overall rise-to-fall movement occurs across time',
            'rise_rate_CURVE': 'Rate of TAC increase from curve threshold to peak (µg/(L·h))',
            'rise_rate_point_to_point_CURVE': 'Average rate of ascending point-to-point TAC changes (µg/(L·min))',
            'rise_duration_point_to_point_CURVE': 'Total duration of ascending point-to-point TAC changes (hours)',
            'rise_rate_1hr_CURVE': 'Rate of TAC increase over first hour, bounded by peak (µg/(L·h))',
            'rise_rate_2hr_CURVE': 'Rate of TAC increase over first 2 hours, bounded by peak (µg/(L·h))',
            'fall_rate_CURVE': 'Rate of TAC decrease from peak to curve end (µg/(L·h))',
            'fall_rate_point_to_point_CURVE': 'Average rate of descending point-to-point TAC changes (µg/(L·min))',
            'fall_duration_point_to_point_CURVE': 'Total duration of descending point-to-point TAC changes (hours)',
            'fall_rate_1hr_CURVE': 'Rate of TAC decrease over last hour, bounded by peak (µg/(L·h))',
            'fall_rate_2hr_CURVE': 'Rate of TAC decrease over last 2 hours, bounded by peak (µg/(L·h))',
            'rise_duration_CURVE': 'Duration of TAC increase (hours)',
            'fall_duration_CURVE': 'Duration of TAC decrease (hours)',
            
            # Quality features
            'total_duration_CURVE': 'Total duration of the region (hours)',
            'device_turned_on_duration_CURVE': 'Duration device was turned on (hours)',
            'device_turned_on_percent_CURVE': 'Percentage of time device was turned on',
            'device_worn_duration_CURVE': 'Duration device was worn (hours)',
            'device_worn_percent_CURVE': 'Percentage of time device was worn',
            'flatline_max_CURVE': 'Maximum duration of flatline (hours)',
            'flatlined_percent_CURVE': 'Percentage of flatlined data',
            'imputed_jump_duration_CURVE': 'Duration of imputed jump data (hours)',
            'imputed_jump_percent_CURVE': 'Percentage of imputed jump data',
            'unimputed_jump_duration_CURVE': 'Duration of unimputed jump data (hours)',
            'unimputed_jump_percent_CURVE': 'Percentage of unimputed jump data',
            'total_jump_duration_CURVE': 'Total duration of jump data (hours)',
            'total_jump_percent_CURVE': 'Total percentage of jump data',
            'total_jump_duration_inclusive_CURVE': 'Total duration of jump data (hours); inclusive: jump==1 only, no mutual exclusivity',
            'total_jump_percent_inclusive_CURVE': 'Total percentage of jump data; inclusive: jump==1 only, no mutual exclusivity',
            'jump_imputation_ratio_CURVE': 'Ratio of jump data that has been imputed (0-1)',
            'total_gap_duration_inclusive_CURVE': 'Total duration of gap data (hours); inclusive: gap==1 only, no mutual exclusivity',
            'total_gap_percent_inclusive_CURVE': 'Total percentage of gap data; inclusive: gap==1 only, no mutual exclusivity',
            'total_non_wear_duration_inclusive_CURVE': 'Total duration of non-wear data (hours); inclusive: non_wear==1 only, no mutual exclusivity',
            'total_non_wear_percent_inclusive_CURVE': 'Total percentage of non-wear data; inclusive: non_wear==1 only, no mutual exclusivity',
            'total_plummet_duration_inclusive_CURVE': 'Total duration of plummet data (hours); inclusive: plummet==1 only, no mutual exclusivity',
            'total_plummet_percent_inclusive_CURVE': 'Total percentage of plummet data; inclusive: plummet==1 only, no mutual exclusivity',
            'total_extreme_negative_duration_inclusive_CURVE': 'Total duration of extreme negative data (hours); inclusive: extreme_negative==1 only, no mutual exclusivity',
            'total_extreme_negative_percent_inclusive_CURVE': 'Total percentage of extreme negative data; inclusive: extreme_negative==1 only, no mutual exclusivity',
            'imputed_duration_CURVE': 'Duration of imputed data (hours)',
            'imputed_percent_CURVE': 'Percentage of imputed data',
            'total_low_quality_duration_CURVE': 'Duration of low quality data (hours)',
            'total_low_quality_percent_CURVE': 'Percentage of low quality data',
            'high_quality_duration_CURVE': 'Duration of high-quality data in the curve (hours)',
            'high_quality_percent_CURVE': 'Percentage of high-quality data in the curve',
            'high_quality_above_threshold_duration_CURVE': 'Duration of high-quality TAC values at or above the curve threshold (hours)',

            # Imputation ratio features
            'jump_imputation_ratio_CURVE': 'Ratio of jump data that has been imputed (0-1)',
            'plummet_imputation_ratio_CURVE': 'Ratio of plummet data that has been imputed (0-1)',
            'extreme_negative_imputation_ratio_CURVE': 'Ratio of extreme negative data that has been imputed (0-1)',
            'gap_imputation_ratio_CURVE': 'Ratio of gap data that has been imputed (0-1)',
            'non_wear_imputation_ratio_CURVE': 'Ratio of non-wear data that has been imputed (0-1)',
            'low_quality_imputation_ratio_CURVE': 'Ratio of overall low quality data that has been imputed (0-1)',
            
            # Categorical features
            'CURVE_VALID': 'Whether the curve meets validity criteria (1=valid)',
            'REGION_VALID': 'Whether both the curve and periphery meet validity criteria (1=valid)',
            'device_count_REGION': 'Number of devices in the region'
        }

        # Person-level feature types
        self._person_level_feature_types = {
            # Continuous TAC features
            'duration_CURVE': 'numeric',
            'auc_total_CURVE': 'numeric',
            'auc_relative_CURVE': 'numeric',
            'peak_CURVE': 'numeric',
            'relative_peak_CURVE': 'numeric',
            'rise_fall_rate_CURVE': 'numeric',
            'rise_rate_CURVE': 'numeric',
            'rise_rate_point_to_point_CURVE': 'numeric',
            'rise_duration_point_to_point_CURVE': 'numeric',
            'rise_rate_1hr_CURVE': 'numeric',
            'rise_rate_2hr_CURVE': 'numeric',
            'fall_rate_CURVE': 'numeric',
            'fall_rate_point_to_point_CURVE': 'numeric',
            'fall_duration_point_to_point_CURVE': 'numeric',
            'fall_rate_1hr_CURVE': 'numeric',
            'fall_rate_2hr_CURVE': 'numeric',
            'rise_duration_CURVE': 'numeric',
            'fall_duration_CURVE': 'numeric',
            'ascending_imputed_percent_CURVE': 'numeric',
            'descending_imputed_percent_CURVE': 'numeric',
            
            # Quality features
            'total_duration_CURVE': 'numeric',
            'device_turned_on_duration_CURVE': 'numeric',
            'device_turned_on_percent_CURVE': 'numeric',
            'device_worn_duration_CURVE': 'numeric',
            'device_worn_percent_CURVE': 'numeric',
            'flatline_max_CURVE': 'numeric',
            'flatlined_percent_CURVE': 'numeric',
            'imputed_jump_duration_CURVE': 'numeric',
            'imputed_jump_percent_CURVE': 'numeric',
            'total_jump_duration_inclusive_CURVE': 'numeric',
            'total_jump_percent_inclusive_CURVE': 'numeric',
            'total_gap_duration_inclusive_CURVE': 'numeric',
            'total_gap_percent_inclusive_CURVE': 'numeric',
            'total_non_wear_duration_inclusive_CURVE': 'numeric',
            'total_non_wear_percent_inclusive_CURVE': 'numeric',
            'total_plummet_duration_inclusive_CURVE': 'numeric',
            'total_plummet_percent_inclusive_CURVE': 'numeric',
            'total_extreme_negative_duration_inclusive_CURVE': 'numeric',
            'total_extreme_negative_percent_inclusive_CURVE': 'numeric',
            'imputed_duration_CURVE': 'numeric',
            'imputed_percent_CURVE': 'numeric',
            'total_low_quality_duration_CURVE': 'numeric',
            'total_low_quality_percent_CURVE': 'numeric',
            'high_quality_duration_CURVE': 'numeric',
            'high_quality_percent_CURVE': 'numeric',
            'high_quality_above_threshold_duration_CURVE': 'numeric',

            # Imputation ratio features
            'jump_imputation_ratio_CURVE': 'numeric',
            'plummet_imputation_ratio_CURVE': 'numeric',
            'extreme_negative_imputation_ratio_CURVE': 'numeric',
            'gap_imputation_ratio_CURVE': 'numeric',
            'non_wear_imputation_ratio_CURVE': 'numeric',
            'low_quality_imputation_ratio_CURVE': 'numeric',
            
            # Negative value features
            'sub_negative_10_sum_CURVE': 'numeric',
            
            # Categorical features
            'CURVE_VALID': 'categorical',
            'PERIPHERY_VALID': 'categorical',
            'REGION_VALID': 'categorical',
            'device_count_REGION': 'categorical',

            # Region quality features
            'started_curve_count_REGION': 'numeric',
            'complete_curve_count_REGION': 'numeric',
            'total_duration_REGION': 'numeric',
            'device_turned_on_duration_REGION': 'numeric',
            'device_turned_on_percent_REGION': 'numeric',
            'device_worn_duration_REGION': 'numeric',
            'device_worn_percent_REGION': 'numeric',
            'flatline_max_REGION': 'numeric',
            'flatlined_percent_REGION': 'numeric',
            'imputed_duration_REGION': 'numeric',
            'imputed_percent_REGION': 'numeric',
            'imputed_low_quality_duration_REGION': 'numeric',
            'imputed_low_quality_percent_REGION': 'numeric',
            'unimputed_low_quality_duration_REGION': 'numeric',
            'unimputed_low_quality_percent_REGION': 'numeric',
            'total_low_quality_duration_REGION': 'numeric',
            'total_low_quality_percent_REGION': 'numeric',
            'imputed_gap_duration_REGION': 'numeric',
            'imputed_gap_percent_REGION': 'numeric',
            'unimputed_gap_duration_REGION': 'numeric',
            'unimputed_gap_percent_REGION': 'numeric',
            'total_gap_duration_REGION': 'numeric',
            'total_gap_percent_REGION': 'numeric',
            'total_gap_duration_inclusive_REGION': 'numeric',
            'total_gap_percent_inclusive_REGION': 'numeric',
            'gap_imputation_ratio_REGION': 'numeric',
            'imputed_non_wear_duration_REGION': 'numeric',
            'imputed_non_wear_percent_REGION': 'numeric',
            'unimputed_non_wear_duration_REGION': 'numeric',
            'unimputed_non_wear_percent_REGION': 'numeric',
            'total_non_wear_duration_REGION': 'numeric',
            'total_non_wear_percent_REGION': 'numeric',
            'total_non_wear_duration_inclusive_REGION': 'numeric',
            'total_non_wear_percent_inclusive_REGION': 'numeric',
            'non_wear_imputation_ratio_REGION': 'numeric',
            'imputed_jump_duration_REGION': 'numeric',
            'imputed_jump_percent_REGION': 'numeric',
            'unimputed_jump_duration_REGION': 'numeric',
            'unimputed_jump_percent_REGION': 'numeric',
            'total_jump_duration_REGION': 'numeric',
            'total_jump_percent_REGION': 'numeric',
            'total_jump_duration_inclusive_REGION': 'numeric',
            'total_jump_percent_inclusive_REGION': 'numeric',
            'jump_imputation_ratio_REGION': 'numeric',
            'imputed_plummet_duration_REGION': 'numeric',
            'imputed_plummet_percent_REGION': 'numeric',
            'unimputed_plummet_duration_REGION': 'numeric',
            'unimputed_plummet_percent_REGION': 'numeric',
            'total_plummet_duration_REGION': 'numeric',
            'total_plummet_percent_REGION': 'numeric',
            'total_plummet_duration_inclusive_REGION': 'numeric',
            'total_plummet_percent_inclusive_REGION': 'numeric',
            'plummet_imputation_ratio_REGION': 'numeric',
            'imputed_extreme_negative_duration_REGION': 'numeric',
            'imputed_extreme_negative_percent_REGION': 'numeric',
            'unimputed_extreme_negative_duration_REGION': 'numeric',
            'unimputed_extreme_negative_percent_REGION': 'numeric',
            'total_extreme_negative_duration_REGION': 'numeric',
            'total_extreme_negative_percent_REGION': 'numeric',
            'total_extreme_negative_duration_inclusive_REGION': 'numeric',
            'total_extreme_negative_percent_inclusive_REGION': 'numeric',
            'sub_negative_10_sum_REGION': 'numeric',
            'extreme_negative_imputation_ratio_REGION': 'numeric',
            'low_quality_imputation_ratio_REGION': 'numeric',
            'total_gaps_and_non_wear_percent_REGION': 'numeric',
        }

        # Periphery Before feature types
        self._periphery_before_feature_types = {
            'total_duration_PERIPHERY_BEFORE': 'numeric',
            'device_turned_on_duration_PERIPHERY_BEFORE': 'numeric',
            'device_turned_on_percent_PERIPHERY_BEFORE': 'numeric',
            'device_worn_duration_PERIPHERY_BEFORE': 'numeric',
            'device_worn_percent_PERIPHERY_BEFORE': 'numeric',
            'imputed_duration_PERIPHERY_BEFORE': 'numeric',
            'imputed_percent_PERIPHERY_BEFORE': 'numeric',
            'imputed_low_quality_duration_PERIPHERY_BEFORE': 'numeric',
            'imputed_low_quality_percent_PERIPHERY_BEFORE': 'numeric',
            'unimputed_low_quality_duration_PERIPHERY_BEFORE': 'numeric',
            'unimputed_low_quality_percent_PERIPHERY_BEFORE': 'numeric',
            'total_low_quality_duration_PERIPHERY_BEFORE': 'numeric',
            'total_low_quality_percent_PERIPHERY_BEFORE': 'numeric',
            'imputed_gap_duration_PERIPHERY_BEFORE': 'numeric',
            'imputed_gap_percent_PERIPHERY_BEFORE': 'numeric',
            'unimputed_gap_duration_PERIPHERY_BEFORE': 'numeric',
            'unimputed_gap_percent_PERIPHERY_BEFORE': 'numeric',
            'total_gap_duration_PERIPHERY_BEFORE': 'numeric',
            'total_gap_percent_PERIPHERY_BEFORE': 'numeric',
            'total_gap_duration_inclusive_PERIPHERY_BEFORE': 'numeric',
            'total_gap_percent_inclusive_PERIPHERY_BEFORE': 'numeric',
            'gap_imputation_ratio_PERIPHERY_BEFORE': 'numeric',
            'imputed_non_wear_duration_PERIPHERY_BEFORE': 'numeric',
            'imputed_non_wear_percent_PERIPHERY_BEFORE': 'numeric',
            'unimputed_non_wear_duration_PERIPHERY_BEFORE': 'numeric',
            'unimputed_non_wear_percent_PERIPHERY_BEFORE': 'numeric',
            'total_non_wear_duration_PERIPHERY_BEFORE': 'numeric',
            'total_non_wear_percent_PERIPHERY_BEFORE': 'numeric',
            'total_non_wear_duration_inclusive_PERIPHERY_BEFORE': 'numeric',
            'total_non_wear_percent_inclusive_PERIPHERY_BEFORE': 'numeric',
            'non_wear_imputation_ratio_PERIPHERY_BEFORE': 'numeric',
            'imputed_jump_duration_PERIPHERY_BEFORE': 'numeric',
            'imputed_jump_percent_PERIPHERY_BEFORE': 'numeric',
            'unimputed_jump_duration_PERIPHERY_BEFORE': 'numeric',
            'unimputed_jump_percent_PERIPHERY_BEFORE': 'numeric',
            'total_jump_duration_PERIPHERY_BEFORE': 'numeric',
            'total_jump_percent_PERIPHERY_BEFORE': 'numeric',
            'total_jump_duration_inclusive_PERIPHERY_BEFORE': 'numeric',
            'total_jump_percent_inclusive_PERIPHERY_BEFORE': 'numeric',
            'jump_imputation_ratio_PERIPHERY_BEFORE': 'numeric',
            'imputed_plummet_duration_PERIPHERY_BEFORE': 'numeric',
            'imputed_plummet_percent_PERIPHERY_BEFORE': 'numeric',
            'unimputed_plummet_duration_PERIPHERY_BEFORE': 'numeric',
            'unimputed_plummet_percent_PERIPHERY_BEFORE': 'numeric',
            'total_plummet_duration_PERIPHERY_BEFORE': 'numeric',
            'total_plummet_percent_PERIPHERY_BEFORE': 'numeric',
            'total_plummet_duration_inclusive_PERIPHERY_BEFORE': 'numeric',
            'total_plummet_percent_inclusive_PERIPHERY_BEFORE': 'numeric',
            'plummet_imputation_ratio_PERIPHERY_BEFORE': 'numeric',
            'imputed_extreme_negative_duration_PERIPHERY_BEFORE': 'numeric',
            'imputed_extreme_negative_percent_PERIPHERY_BEFORE': 'numeric',
            'unimputed_extreme_negative_duration_PERIPHERY_BEFORE': 'numeric',
            'unimputed_extreme_negative_percent_PERIPHERY_BEFORE': 'numeric',
            'total_extreme_negative_duration_PERIPHERY_BEFORE': 'numeric',
            'total_extreme_negative_percent_PERIPHERY_BEFORE': 'numeric',
            'total_extreme_negative_duration_inclusive_PERIPHERY_BEFORE': 'numeric',
            'total_extreme_negative_percent_inclusive_PERIPHERY_BEFORE': 'numeric',
            'sub_negative_10_sum_PERIPHERY_BEFORE': 'numeric',
            'extreme_negative_imputation_ratio_PERIPHERY_BEFORE': 'numeric',
            'low_quality_imputation_ratio_PERIPHERY_BEFORE': 'numeric',
            'total_gaps_and_non_wear_percent_PERIPHERY_BEFORE': 'numeric',
        }

        # Periphery After feature types
        self._periphery_after_feature_types = {
            'total_duration_PERIPHERY_AFTER': 'numeric',
            'device_turned_on_duration_PERIPHERY_AFTER': 'numeric',
            'device_turned_on_percent_PERIPHERY_AFTER': 'numeric',
            'device_worn_duration_PERIPHERY_AFTER': 'numeric',
            'device_worn_percent_PERIPHERY_AFTER': 'numeric',
            'imputed_duration_PERIPHERY_AFTER': 'numeric',
            'imputed_percent_PERIPHERY_AFTER': 'numeric',
            'imputed_low_quality_duration_PERIPHERY_AFTER': 'numeric',
            'imputed_low_quality_percent_PERIPHERY_AFTER': 'numeric',
            'unimputed_low_quality_duration_PERIPHERY_AFTER': 'numeric',
            'unimputed_low_quality_percent_PERIPHERY_AFTER': 'numeric',
            'total_low_quality_duration_PERIPHERY_AFTER': 'numeric',
            'total_low_quality_percent_PERIPHERY_AFTER': 'numeric',
            'imputed_gap_duration_PERIPHERY_AFTER': 'numeric',
            'imputed_gap_percent_PERIPHERY_AFTER': 'numeric',
            'unimputed_gap_duration_PERIPHERY_AFTER': 'numeric',
            'unimputed_gap_percent_PERIPHERY_AFTER': 'numeric',
            'total_gap_duration_PERIPHERY_AFTER': 'numeric',
            'total_gap_percent_PERIPHERY_AFTER': 'numeric',
            'total_gap_duration_inclusive_PERIPHERY_AFTER': 'numeric',
            'total_gap_percent_inclusive_PERIPHERY_AFTER': 'numeric',
            'gap_imputation_ratio_PERIPHERY_AFTER': 'numeric',
            'imputed_non_wear_duration_PERIPHERY_AFTER': 'numeric',
            'imputed_non_wear_percent_PERIPHERY_AFTER': 'numeric',
            'unimputed_non_wear_duration_PERIPHERY_AFTER': 'numeric',
            'unimputed_non_wear_percent_PERIPHERY_AFTER': 'numeric',
            'total_non_wear_duration_PERIPHERY_AFTER': 'numeric',
            'total_non_wear_percent_PERIPHERY_AFTER': 'numeric',
            'total_non_wear_duration_inclusive_PERIPHERY_AFTER': 'numeric',
            'total_non_wear_percent_inclusive_PERIPHERY_AFTER': 'numeric',
            'non_wear_imputation_ratio_PERIPHERY_AFTER': 'numeric',
            'imputed_jump_duration_PERIPHERY_AFTER': 'numeric',
            'imputed_jump_percent_PERIPHERY_AFTER': 'numeric',
            'unimputed_jump_duration_PERIPHERY_AFTER': 'numeric',
            'unimputed_jump_percent_PERIPHERY_AFTER': 'numeric',
            'total_jump_duration_PERIPHERY_AFTER': 'numeric',
            'total_jump_percent_PERIPHERY_AFTER': 'numeric',
            'total_jump_duration_inclusive_PERIPHERY_AFTER': 'numeric',
            'total_jump_percent_inclusive_PERIPHERY_AFTER': 'numeric',
            'jump_imputation_ratio_PERIPHERY_AFTER': 'numeric',
            'imputed_plummet_duration_PERIPHERY_AFTER': 'numeric',
            'imputed_plummet_percent_PERIPHERY_AFTER': 'numeric',
            'unimputed_plummet_duration_PERIPHERY_AFTER': 'numeric',
            'unimputed_plummet_percent_PERIPHERY_AFTER': 'numeric',
            'total_plummet_duration_PERIPHERY_AFTER': 'numeric',
            'total_plummet_percent_PERIPHERY_AFTER': 'numeric',
            'total_plummet_duration_inclusive_PERIPHERY_AFTER': 'numeric',
            'total_plummet_percent_inclusive_PERIPHERY_AFTER': 'numeric',
            'plummet_imputation_ratio_PERIPHERY_AFTER': 'numeric',
            'imputed_extreme_negative_duration_PERIPHERY_AFTER': 'numeric',
            'imputed_extreme_negative_percent_PERIPHERY_AFTER': 'numeric',
            'unimputed_extreme_negative_duration_PERIPHERY_AFTER': 'numeric',
            'unimputed_extreme_negative_percent_PERIPHERY_AFTER': 'numeric',
            'total_extreme_negative_duration_PERIPHERY_AFTER': 'numeric',
            'total_extreme_negative_percent_PERIPHERY_AFTER': 'numeric',
            'total_extreme_negative_duration_inclusive_PERIPHERY_AFTER': 'numeric',
            'total_extreme_negative_percent_inclusive_PERIPHERY_AFTER': 'numeric',
            'sub_negative_10_sum_PERIPHERY_AFTER': 'numeric',
            'extreme_negative_imputation_ratio_PERIPHERY_AFTER': 'numeric',
            'low_quality_imputation_ratio_PERIPHERY_AFTER': 'numeric',
            'total_gaps_and_non_wear_percent_PERIPHERY_AFTER': 'numeric',
        }

        # Imputation feature descriptions (used in Imputations tab)
        self._imputation_feature_descriptions = {
            'region_start': 'Start time of the imputed region',
            'region_end': 'End time of the imputed region',
            'region_length': 'Length (duration) of the imputed region (minutes or hours)',
            'worn_minutes_before': 'Minutes device was worn before the imputed region',
            'worn_percent_before': 'Percent of time device was worn before the imputed region',
            'worn_minutes_after': 'Minutes device was worn after the imputed region',
            'worn_percent_after': 'Percent of time device was worn after the imputed region',
            'min_training_data_required': 'Minimum training data required for imputation',
            'was_imputed': 'Whether this region was imputed (1=True, 0=False)',
            'reason_not_imputed': 'Reason why the region was not imputed (if applicable)',
            'high_quality_before': 'High quality data points before the imputed region',
            'low_quality_before': 'Low quality data points before the imputed region',
            'high_quality_after': 'High quality data points after the imputed region',
            'low_quality_after': 'Low quality data points after the imputed region',
            'total_training_before': 'Total training data points before the imputed region',
            'total_training_after': 'Total training data points after the imputed region',
        }

        # Day feature descriptions (used in day-level analysis)
        self._day_feature_descriptions = {
            # Basic identifiers
            'SubID': 'Subject identifier',
            'Dataset_ID': 'Dataset identifier',
            'day_no': 'Day number (sequential day in the study)',
            'begin_day': 'Start time of the day (social day boundary)',
            'day_of_week_cat': (
                'Calendar weekday derived from ``begin_day``: integer **1 = Monday** through **7 = Sunday** '
                '(pandas weekday convention); missing when ``begin_day`` is missing or invalid'
            ),
            'new_device': '1 if device_two is non-null (two distinct devices contributed data that day, indicating a mid-burst device swap); 0 otherwise',
            'end_day': 'End time of the day (social day boundary)',
            
            # Device information
            'device_ids': 'List of device IDs present during this day',
            'device_one': 'Primary device ID',
            'device_two': 'Secondary device ID (if present)',
            'device_count': 'Number of devices present during this day',
            'firmware': 'Firmware version of the device(s)',
            
            # Day-level timing
            'day_hours': 'Total hours in the day (always 24.0 for complete days)',
            'device_turned_on_duration': 'Duration device was turned on (hours)',
            'device_turned_on_percentage_of_day': 'Percentage of day device was turned on',
            
            # Device wear information
            'device_worn_duration': 'Duration device was worn (hours) [using algorithm]',
            'device_worn_percent_of_device_on': 'Percentage of device-on time that device was worn',
            'device_worn_percent_of_day': 'Percentage of day that device was worn',

            # Day-level TAC summaries (minute rows in the social day); exploratory — coverage varies; do not rely on these as primary inputs yet
            'tac_day_mean': (
                'Mean TAC (µg/L) across minute rows in the day window (**coerced numeric**; NaNs excluded from the mean)'
            ),
            'tac_day_median': 'Median TAC (µg/L) across minute rows in the day window',
            'tac_day_sd': 'Sample standard deviation of TAC across minute rows (pandas ``std``, ddof=1; NaN if insufficient finite values)',
            'tac_day_min': 'Minimum TAC (µg/L) among minute rows in the day window',
            'tac_day_max': 'Maximum TAC (µg/L) among minute rows in the day window',
            'tac_day_q25': '25th percentile of TAC across minute rows in the day window',
            'tac_day_q75': '75th percentile of TAC across minute rows in the day window',
            'tac_non_missing_minute_count': (
                'Count of minute rows in the day window with non-null TAC after numeric coercion'
            ),
            'tac_non_missing_fraction_of_rows': (
                '``tac_non_missing_minute_count`` divided by the number of minute rows in the day segment (0 when empty)'
            ),

            # Time-region TAC features — Q1 (hours 0–6) and Q2_Q4 (hours 6–24) of the social day
            'tac_q1_mean': 'Mean TAC (µg/L) across minute rows in Q1 (hours 0–6 of social day); NaN if no valid TAC data',
            'tac_q1_max': 'Maximum TAC (µg/L) among minute rows in Q1; NaN if no valid TAC data',
            'tac_q1_above_threshold_percent': 'Fraction of minute rows in Q1 where TAC ≥ curve_threshold (0–1); NaN if no data',
            'tac_q1_auc': 'Area under the TAC curve in Q1 (trapezoidal rule, dx=1 minute → µg/L·min); NaN if no valid TAC data',
            'tac_q1_rise_rate_point_to_point': 'Mean ascending point-to-point TAC rate within Q1; 0 if no ascending steps; NaN if no data',
            'tac_q1_fall_rate_point_to_point': 'Mean descending point-to-point TAC rate within Q1; 0 if no descending steps; NaN if no data',
            'tac_q1_ascending_duration': 'Total duration (hours) of ascending TAC steps within Q1; 0 if none; NaN if no data',
            'tac_q1_descending_duration': 'Total duration (hours) of descending TAC steps within Q1; 0 if none; NaN if no data',

            'tac_q2_q4_mean': 'Mean TAC (µg/L) across minute rows in Q2–Q4 (hours 6–24 of social day); NaN if no valid TAC data',
            'tac_q2_q4_max': 'Maximum TAC (µg/L) among minute rows in Q2–Q4; NaN if no valid TAC data',
            'tac_q2_q4_above_threshold_percent': 'Fraction of minute rows in Q2–Q4 where TAC ≥ curve_threshold (0–1); NaN if no data',
            'tac_q2_q4_auc': 'Area under the TAC curve in Q2–Q4 (trapezoidal rule, dx=1 minute → µg/L·min); NaN if no valid TAC data',
            'tac_q2_q4_rise_rate_point_to_point': 'Mean ascending point-to-point TAC rate within Q2–Q4; 0 if no ascending steps; NaN if no data',
            'tac_q2_q4_fall_rate_point_to_point': 'Mean descending point-to-point TAC rate within Q2–Q4; 0 if no descending steps; NaN if no data',
            'tac_q2_q4_ascending_duration': 'Total duration (hours) of ascending TAC steps within Q2–Q4; 0 if none; NaN if no data',
            'tac_q2_q4_descending_duration': 'Total duration (hours) of descending TAC steps within Q2–Q4; 0 if none; NaN if no data',

            # Burst prior medians across all qualifying curves on prior days within the burst
            'median_prior_peak_CURVE': (
                'Median ``peak_CURVE`` across all qualifying curves on prior burst days '
                '(HQ% >= 75% and HQ duration >= 1 hour); NaN on the first day or if no qualifying prior curves'
            ),
            'median_prior_auc_total_CURVE': (
                'Median ``auc_total_CURVE`` across all qualifying prior curves in the burst; NaN first day'
            ),
            'median_prior_duration_CURVE': (
                'Median ``duration_CURVE`` across all qualifying prior curves in the burst; NaN first day'
            ),
            'median_prior_total_low_quality_percent_CURVE': (
                'Median ``total_low_quality_percent_CURVE`` across all qualifying prior curves in the burst; '
                'NaN first day'
            ),
            'median_prior_rise_rate_CURVE': (
                'Median ``rise_rate_CURVE`` across all qualifying prior curves in the burst; NaN first day'
            ),
            'median_prior_fall_rate_CURVE': (
                'Median ``fall_rate_CURVE`` across all qualifying prior curves in the burst; NaN first day'
            ),
            'median_prior_rise_rate_point_to_point_CURVE': (
                'Median ``rise_rate_point_to_point_CURVE`` across all qualifying prior curves in the burst; NaN first day'
            ),
            'median_prior_fall_rate_point_to_point_CURVE': (
                'Median ``fall_rate_point_to_point_CURVE`` across all qualifying prior curves in the burst; NaN first day'
            ),
            'median_prior_fall_duration_CURVE': (
                'Median ``fall_duration_CURVE`` across all qualifying prior curves in the burst; NaN first day'
            ),
            'median_prior_rise_duration_CURVE': (
                'Median ``rise_duration_CURVE`` across all qualifying prior curves in the burst; NaN first day'
            ),

            # Prior burst-day medians of day-level TAC region features
            'median_prior_tac_q1_mean': 'Expanding median of tac_q1_mean across prior days in the burst; NaN first day',
            'median_prior_tac_q1_max': 'Expanding median of tac_q1_max across prior days in the burst; NaN first day',
            'median_prior_tac_q1_auc': 'Expanding median of tac_q1_auc across prior days in the burst; NaN first day',
            'median_prior_tac_q2_q4_mean': 'Expanding median of tac_q2_q4_mean across prior days in the burst; NaN first day',
            'median_prior_tac_q2_q4_max': 'Expanding median of tac_q2_q4_max across prior days in the burst; NaN first day',
            'median_prior_tac_q2_q4_auc': 'Expanding median of tac_q2_q4_auc across prior days in the burst; NaN first day',
            'median_prior_curve_1_start_offset_hours': 'Expanding median of curve_1_start_offset_hours across prior days in the burst; captures typical top-curve start time; NaN first day',
            'median_prior_curve_1_end_offset_hours': 'Expanding median of curve_1_end_offset_hours across prior days in the burst; captures typical top-curve end time; NaN first day',

            # Per-slot derived features (slots ordered by descending high_quality_duration_CURVE)
            'curve_{n}_overlap_fraction': (
                '``curve_{n}_overlap_hours`` divided by ``curve_{n}_duration_CURVE`` (fraction of the curve that '
                'falls within the social day); NaN if duration missing or <= 0'
            ),
            'curve_{n}_start_offset_hours': (
                'Hours from ``begin_day`` to ``begin_CURVE`` for the curve in slot ``n``; '
                'negative if the curve started before the social day (prior-day spillover); 0 at day start'
            ),
            'curve_{n}_end_offset_hours': (
                'Hours from ``begin_day`` to ``end_CURVE`` for the curve in slot ``n``; '
                '>24 if the curve extends past the social day; values near 24 = curve ended at day end'
            ),
            'device_worn_duration_cutoff': 'Duration device was worn [using 28 Celcius cutoff]',
            'device_worn_cutoff_percent_of_device_on': 'Percentage of device-on time that device was worn (with cutoff)',
            'device_worn_cutoff_percent_of_day': 'Percentage of day that device was worn (with cutoff)',
            
            # Data quality and imputation
            'imputed_duration': 'Duration of imputed data (hours)',
            'imputed_percent': 'Percentage of day that was imputed',
            
            # Low quality data
            'low_quality_duration': 'Duration of low quality data (hours)',
            'low_quality_percent': 'Percentage of day that was low quality',
            'unimputed_low_quality_duration': 'Duration of unimputed low quality data (hours)',
            'unimputed_low_quality_percent': 'Percentage of day that was unimputed low quality',
            'imputed_low_quality_duration': 'Duration of imputed low quality data (hours)',
            'imputed_low_quality_percent': 'Percentage of day that was imputed low quality',
            'low_quality_percent_of_day': 'Low-quality duration divided by 24 hours (0–1; fixed 24h denominator)',
            'unimputed_low_quality_percent_of_day': 'Unimputed low-quality duration divided by 24 hours (0–1; fixed 24h denominator)',
            'imputed_low_quality_percent_of_day': 'Imputed low-quality duration divided by 24 hours (0–1; fixed 24h denominator)',
            'tac_day_missing': '1 if this day row was inserted to represent a study day with no TAC day-level data; 0 if observed from TAC processing',
            'unknown_reason': 'Reason label for unknown-classified day rows (blank otherwise): prediction_unreliable or missing_day_expanded',
            
            # Gap analysis
            'gap_duration': 'Duration of data gaps (hours)',
            'gap_percent': 'Percentage of day that had gaps',
            'imputed_gap_duration': 'Duration of imputed gaps (hours)',
            'imputed_gap_percent': 'Percentage of day that had imputed gaps',
            'unimputed_gap_duration': 'Duration of unimputed gaps (hours)',
            'unimputed_gap_percent': 'Percentage of day that had unimputed gaps',
            'gap_imputation_ratio': 'Ratio of gaps that were imputed (0-1)',
            
            # Non-wear analysis
            'non_wear_duration': 'Duration of non-wear periods (hours)',
            'non_wear_percent': 'Percentage of day that was non-wear',
            'imputed_non_wear_duration': 'Duration of imputed non-wear (hours)',
            'imputed_non_wear_percent': 'Percentage of day that was imputed non-wear',
            'unimputed_non_wear_duration': 'Duration of unimputed non-wear (hours)',
            'unimputed_non_wear_percent': 'Percentage of day that was unimputed non-wear',
            'non_wear_imputation_ratio': 'Ratio of non-wear that was imputed (0-1)',
            
            # Jump analysis
            'jump_duration': 'Duration of jump artifacts (hours)',
            'jump_percent': 'Percentage of day that had jump artifacts',
            'imputed_jump_duration': 'Duration of imputed jump artifacts (hours)',
            'imputed_jump_percent': 'Percentage of day that had imputed jump artifacts',
            'unimputed_jump_duration': 'Duration of unimputed jump artifacts (hours)',
            'unimputed_jump_percent': 'Percentage of day that had unimputed jump artifacts',
            'jump_imputation_ratio': 'Ratio of jump artifacts that were imputed (0-1)',
            
            # Plummet analysis
            'plummet_duration': 'Duration of plummet artifacts (hours)',
            'plummet_percent': 'Percentage of day that had plummet artifacts',
            'imputed_plummet_duration': 'Duration of imputed plummet artifacts (hours)',
            'imputed_plummet_percent': 'Percentage of day that had imputed plummet artifacts',
            'unimputed_plummet_duration': 'Duration of unimputed plummet artifacts (hours)',
            'unimputed_plummet_percent': 'Percentage of day that had unimputed plummet artifacts',
            'plummet_imputation_ratio': 'Ratio of plummet artifacts that were imputed (0-1)',
            
            # Negative value analysis
            'negative_duration': 'Duration of negative TAC values (hours)',
            'negative_percent': 'Percentage of day that had negative TAC values',
            'extreme_negative_duration': 'Duration of extreme negative TAC values (hours)',
            'extreme_negative_percent': 'Percentage of day that had extreme negative TAC values',
            'extreme_negative_sum': 'Sum of extreme negative TAC values',
            'imputed_extreme_negative_duration': 'Duration of imputed extreme negative values (hours)',
            'imputed_extreme_negative_percent': 'Percentage of day that had imputed extreme negative values',
            'unimputed_extreme_negative_duration': 'Duration of unimputed extreme negative values (hours)',
            'unimputed_extreme_negative_percent': 'Percentage of day that had unimputed extreme negative values',
            'extreme_negative_imputation_ratio': 'Ratio of extreme negative values that were imputed (0-1)',
            
            # Overall quality metrics
            'low_quality_imputation_ratio': 'Overall ratio of low quality data that was imputed (0-1)',
            'total_gaps_and_non_wear_percent': 'Total percentage of gaps and non-wear combined',
            'below_threshold_percent': 'Percentage of data below detection threshold',
            'above_threshold_duration': 'Duration of TAC at or above curve threshold (hours)',
            'above_threshold_percent_of_day': 'Fraction of the 24-hour day with TAC at or above curve threshold (0-1)',
            'above_threshold_high_quality_percent': 'Of above-threshold minutes, fraction that are high-quality (0-1); None if no above-threshold data',
            'above_threshold_high_quality_percent_of_day': (
                'Fraction of the 24-hour day with TAC both at/above curve threshold **and** high-quality '
                '(minutes above threshold & not low-quality, divided by 24 hours; 0 when no data)'
            ),
            
            # Signal characteristics
            'flatline_max': 'Maximum duration of flatline periods (minutes)',
            'flatlined_percent': 'Percentage of day that was flatlined',
            
            # Temperature statistics
            'temp_mean': 'Mean temperature during the day',
            'temp_sd': 'Standard deviation of temperature during the day',
            'temp_min': 'Minimum temperature during the day',
            'temp_max': 'Maximum temperature during the day',
            
            # Motion statistics
            'motion_mean': 'Mean motion during the day',
            'motion_sd': 'Standard deviation of motion during the day',
            'motion_min': 'Minimum motion during the day',
            'motion_max': 'Maximum motion during the day',
            
            # Date information
            'date': 'Date of the day',
            
            # Drinking curve overlap detection
            'drinking_curve_overlap': 'Whether any drinking curve overlapped with this day (0/1)',
            'pred_drinking_day': (
                '08.05 XGBoost drinking-day prediction (1=drinking, 0=not drinking). '
                'Scored for every day, including days with no TAC.'
            ),
            'proba_drinking_day': 'Predicted probability of a drinking day from the 08.05 XGBoost model (0-1).',
            'flag_non_wear': (
                '1 when hours not worn are at least 6 (configurable). Hours not worn = day_hours - device_worn_duration, '
                'which includes powered-off or missing time and on-device non-wear readings.'
            ),
            'flag_very_negative_tac': '1 when extreme_negative_duration (TAC < -15) is at least 1 hour (configurable).',
            'flag_artifacts': '1 when jump_duration + plummet_duration is at least 1 hour (configurable).',
            'flag_uncertain_probability': '1 when proba_drinking_day is inside [0.05, 0.95] (configurable, inclusive).',
            'prediction_reliable': '1 when flag_non_wear, flag_very_negative_tac, flag_artifacts, and flag_uncertain_probability are all 0; else 0.',
            'total_curve_overlap_hours': 'Total hours of curve overlap with this day (sum across all curves)',
            
            # Individual curve overlap details (dynamic based on max curves per day)
            'curve_1_id': 'ID of first curve that overlapped with this day',
            'curve_1_overlap_hours': 'Hours of overlap between first curve and this day',
            'curve_1_duration_CURVE': (
                'Total duration (hours) of the first overlapping curve from ``duration_CURVE``; used with '
                '``curve_1_overlap_hours`` for fraction-on-day when selecting max-HQ attachment'
            ),
            'curve_1_high_quality_duration': (
                'High-quality duration (hours) for **overlap slot 1** (chronologically first overlapping curve), from '
                '``high_quality_duration_CURVE``. Full curve features are available per slot as ``curve_{n}_<feature>``.'
            ),
            'curve_1_extends_prior_day': 'Whether first curve extends into prior day (0/1)',
            'curve_1_extends_next_day': 'Whether first curve extends into next day (0/1)',
            'curve_2_id': 'ID of second curve that overlapped with this day',
            'curve_2_overlap_hours': 'Hours of overlap between second curve and this day',
            'curve_2_duration_CURVE': 'Total duration (hours) of the second overlapping curve (``duration_CURVE``)',
            'curve_2_high_quality_duration': 'High-quality duration from second overlapping curve (carried forward from curve features)',
            'curve_2_extends_prior_day': 'Whether second curve extends into prior day (0/1)',
            'curve_2_extends_next_day': 'Whether second curve extends into next day (0/1)',
            'curve_3_id': 'ID of third curve that overlapped with this day',
            'curve_3_overlap_hours': 'Hours of overlap between third curve and this day',
            'curve_3_duration_CURVE': 'Total duration (hours) of the third overlapping curve (``duration_CURVE``)',
            'curve_3_high_quality_duration': 'High-quality duration from third overlapping curve (carried forward from curve features)',
            'curve_3_extends_prior_day': 'Whether third curve extends into prior day (0/1)',
            'curve_3_extends_next_day': 'Whether third curve extends into next day (0/1)',
            'annotation_drinking_day': (
                'Manual or review annotation for drinking vs non-drinking day (e.g. filled in Excel or '
                'downstream); initialized as blank when building curve-attachment exports.'
            ),

            # Morning EMA merge and TAC agreement (day rows)
            'morning_self_report_alcohol': (
                'Merged morning self-report of alcohol use for this day (0/1 when a response exists; '
                'missing if no morning item matched this day row). Source column is morning.csv '
                'mr_al_y (protocol name mr_alc: Did you use alcohol yesterday?; 1=yes, 2=no). '
                'Not an ML training feature/label.'
            ),
            'morning_report_matched': (
                '1 if a morning alcohol response was linked to this day row; 0 if no match'
            ),
            'morning_merge_date': (
                'Calendar date used when joining morning EMA to this TAC day (calendar merge path)'
            ),
            'mr_numdk': (
                'Morning self-report: total standard drinks yesterday (numeric). Attached for QC only; '
                'not used as an ML feature. CONDITION_SKIPPED coerced to missing.'
            ),
            'mr_alst': (
                'Protocol: time of first sip yesterday. Not present in the current morning report '
                '(column kept as missing on day rows when absent).'
            ),
            'mr_alfn': (
                'Protocol: time of last sip yesterday. Not present in the current morning report '
                '(column kept as missing on day rows when absent).'
            ),
            'mr_altim_1': (
                'Morning alcohol time-bin checkbox 1 (0/1). Export stand-in for protocol mr_alst/mr_alfn. QC only.'
            ),
            'mr_altim_2': (
                'Morning alcohol time-bin checkbox 2 (0/1). QC only; not an ML feature.'
            ),
            'mr_altim_3': (
                'Morning alcohol time-bin checkbox 3 (0/1). QC only; not an ML feature.'
            ),
            'mr_altim_4': (
                'Morning alcohol time-bin checkbox 4 (0/1). QC only; not an ML feature.'
            ),
            'mr_altim_5': (
                'Morning alcohol time-bin checkbox 5 (0/1). QC only; not an ML feature.'
            ),
            'mr_altim_6': (
                'Morning alcohol time-bin checkbox 6 (0/1). QC only; not an ML feature.'
            ),
            'mr_altim_7': (
                'Morning alcohol time-bin checkbox 7 (0/1). QC only; not an ML feature.'
            ),
            'inside_burst': (
                '1 if this day’s calendar date falls inside the skyn wear / study window for that burst; '
                '0 if outside (used to subset Agreement_summary and burst grids)'
            ),
            'self_report_and_tac_comparison': (
                'Morning alcohol self-report vs drinking-day flag (pred_drinking_day, 08.05 XGBoost). '
                'true_positive, false_positive, true_negative, false_negative when the non-drinking day has '
                'prediction_reliable = 1; true_unknown / false_unknown when the non-drinking day has '
                'prediction_reliable = 0 and self-report is present; '
                'positive_without_self_report when drinking day '
                'and no morning SR; negative_without_self_report when no drinking day, no morning SR, and '
                'prediction_reliable = 1; unknown_sr_missing when no drinking day, no morning '
                'SR, and prediction_reliable = 0'
            ),
        }

        # Tab descriptions for different export types
        self._curve_features_tab_descriptions = {
            'Variable Key': 'Descriptions of all variables used in the workbook',
            'Stats': 'Statistical summaries of curve features',
            'Features': 'Curve features and quality metrics for each curve',
            'Person Level Stats': 'Person-level statistics for all curves',
            'Valid Curves': 'Visualization of curves that met validity criteria',
            'Invalid Curves': 'Visualization of curves that did not meet validity criteria',
            'Imputations': 'Information about imputed regions in the data',
            'Run Settings': 'Settings used to generate the curves',
        }

        self._curve_features_with_events_tab_descriptions = {
            'Variable Key': 'Descriptions of all variables used in the workbook',
            'STATS': 'Statistical summaries of curve features and event matching',
            'Features': 'Curve features and quality metrics for each curve',
            'Events': 'Event data and their matching information',
            'Valid Matched by SubID': 'Person-level statistics for valid curves with matched events',
            'Valid Curves': 'Visualization of curves that met validity criteria',
            'Invalid Curves': 'Visualization of curves that did not meet validity criteria',
            'No Curve - EMA Region': 'Visualization of events that did not match to any curve',
            'Imputations': 'Information about imputed regions in the data',
            'Run Settings': 'Settings used to generate the curves and match events',
        }

        # Default TAC and quality features for statistics
        self._stats_features = [
            # TAC curve features
            'duration_CURVE', 'auc_total_CURVE', 'auc_relative_CURVE', 
            'peak_CURVE', 'relative_peak_CURVE',
            'rise_rate_CURVE', 'rise_rate_point_to_point_CURVE', 'rise_rate_1hr_CURVE', 'rise_rate_2hr_CURVE',
            'fall_rate_CURVE', 'fall_rate_point_to_point_CURVE', 'fall_rate_1hr_CURVE', 'fall_rate_2hr_CURVE',
            'fall_duration_CURVE',
            
            # Quality features
            'total_duration_CURVE',
            'device_turned_on_duration_CURVE',
            'device_turned_on_percent_CURVE',
            'device_worn_duration_CURVE',
            'device_worn_percent_CURVE',
            'flatline_max_CURVE',
            'flatlined_percent_CURVE',
            'imputed_jump_duration_CURVE',
            'imputed_jump_percent_CURVE',
            'imputed_duration_CURVE',
            'imputed_percent_CURVE',
            'total_low_quality_duration_CURVE',
            'total_low_quality_percent_CURVE',

            # Negative value features
            'sub_negative_10_sum_CURVE'
        ]

        # Region quality features
        self._region_quality_features = {
            'started_curve_count_REGION': 'Number of discrete curves within the region',
            'complete_curve_count_REGION': 'Number of discrete curves started and ended within the region',
            'total_duration_REGION': 'Total duration of the region (hours)',
            'device_turned_on_duration_REGION': 'Duration device was turned on in region (hours)',
            'device_turned_on_percent_REGION': 'Percentage of time device was turned on in region',
            'device_worn_duration_REGION': 'Duration device was worn in region (hours)',
            'device_worn_percent_REGION': 'Percentage of time device was worn in region',
            'flatline_max_REGION': 'Maximum duration of flatline in region (hours)',
            'flatlined_percent_REGION': 'Percentage of flatlined data in region',
            'imputed_duration_REGION': 'Duration of imputed data in region (hours)',
            'imputed_percent_REGION': 'Percentage of imputed data in region',
            'imputed_low_quality_duration_REGION': 'Duration of imputed low quality data in region (hours)',
            'imputed_low_quality_percent_REGION': 'Percentage of imputed low quality data in region',
            'unimputed_low_quality_duration_REGION': 'Duration of low quality data not imputed in region (hours)',
            'unimputed_low_quality_percent_REGION': 'Percentage of low quality data not imputed in region',
            'total_low_quality_duration_REGION': 'Total duration of low quality data in region (hours)',
            'total_low_quality_percent_REGION': 'Total percentage of low quality data in region',
            'imputed_gap_duration_REGION': 'Duration of imputed gap data in region (hours)',
            'imputed_gap_percent_REGION': 'Percentage of imputed gap data in region',
            'unimputed_gap_duration_REGION': 'Duration of unimputed gap data in region (hours)',
            'unimputed_gap_percent_REGION': 'Percentage of unimputed gap data in region',
            'total_gap_duration_REGION': 'Total duration of gap data in region (hours)',
            'total_gap_percent_REGION': 'Total percentage of gap data in region',
            'total_gap_duration_inclusive_REGION': 'Total duration of gap data in region (hours); inclusive: gap==1 only, no mutual exclusivity',
            'total_gap_percent_inclusive_REGION': 'Total percentage of gap data in region; inclusive: gap==1 only, no mutual exclusivity',
            'gap_imputation_ratio_REGION': 'Ratio of gap data that has been imputed in region (0-1)',
            'imputed_non_wear_duration_REGION': 'Duration of imputed non-wear data in region (hours)',
            'imputed_non_wear_percent_REGION': 'Percentage of imputed non-wear data in region',
            'unimputed_non_wear_duration_REGION': 'Duration of unimputed non-wear data in region (hours)',
            'unimputed_non_wear_percent_REGION': 'Percentage of unimputed non-wear data in region',
            'total_non_wear_duration_REGION': 'Total duration of non-wear data in region (hours)',
            'total_non_wear_percent_REGION': 'Total percentage of non-wear data in region',
            'total_non_wear_duration_inclusive_REGION': 'Total duration of non-wear data in region (hours); inclusive: non_wear==1 only, no mutual exclusivity',
            'total_non_wear_percent_inclusive_REGION': 'Total percentage of non-wear data in region; inclusive: non_wear==1 only, no mutual exclusivity',
            'non_wear_imputation_ratio_REGION': 'Ratio of non-wear data that has been imputed in region (0-1)',
            'imputed_jump_duration_REGION': 'Duration of imputed jump data in region (hours)',
            'imputed_jump_percent_REGION': 'Percentage of imputed jump data in region',
            'unimputed_jump_duration_REGION': 'Duration of unimputed jump data in region (hours)',
            'unimputed_jump_percent_REGION': 'Percentage of unimputed jump data in region',
            'total_jump_duration_REGION': 'Total duration of jump data in region (hours)',
            'total_jump_percent_REGION': 'Total percentage of jump data in region',
            'total_jump_duration_inclusive_REGION': 'Total duration of jump data in region (hours); inclusive: jump==1 only, no mutual exclusivity',
            'total_jump_percent_inclusive_REGION': 'Total percentage of jump data in region; inclusive: jump==1 only, no mutual exclusivity',
            'jump_imputation_ratio_REGION': 'Ratio of jump data that has been imputed in region (0-1)',
            'imputed_plummet_duration_REGION': 'Duration of imputed plummet data in region (hours)',
            'imputed_plummet_percent_REGION': 'Percentage of imputed plummet data in region',
            'unimputed_plummet_duration_REGION': 'Duration of unimputed plummet data in region (hours)',
            'unimputed_plummet_percent_REGION': 'Percentage of unimputed plummet data in region',
            'total_plummet_duration_REGION': 'Total duration of plummet data in region (hours)',
            'total_plummet_percent_REGION': 'Total percentage of plummet data in region',
            'total_plummet_duration_inclusive_REGION': 'Total duration of plummet data in region (hours); inclusive: plummet==1 only, no mutual exclusivity',
            'total_plummet_percent_inclusive_REGION': 'Total percentage of plummet data in region; inclusive: plummet==1 only, no mutual exclusivity',
            'plummet_imputation_ratio_REGION': 'Ratio of plummet data that has been imputed in region (0-1)',
            'imputed_extreme_negative_duration_REGION': 'Duration of imputed extreme negative data in region (hours)',
            'imputed_extreme_negative_percent_REGION': 'Percentage of imputed extreme negative data in region',
            'unimputed_extreme_negative_duration_REGION': 'Duration of unimputed extreme negative data in region (hours)',
            'unimputed_extreme_negative_percent_REGION': 'Percentage of unimputed extreme negative data in region',
            'total_extreme_negative_duration_REGION': 'Total duration of extreme negative data in region (hours)',
            'total_extreme_negative_percent_REGION': 'Total percentage of extreme negative data in region',
            'total_extreme_negative_duration_inclusive_REGION': 'Total duration of extreme negative data in region (hours); inclusive: extreme_negative==1 only, no mutual exclusivity',
            'total_extreme_negative_percent_inclusive_REGION': 'Total percentage of extreme negative data in region; inclusive: extreme_negative==1 only, no mutual exclusivity',
            'sub_negative_10_sum_REGION': 'Sum of all TAC values below -10 in region',
            'extreme_negative_imputation_ratio_REGION': 'Ratio of extreme negative data that has been imputed in region (0-1)',
            'low_quality_imputation_ratio_REGION': 'Ratio of overall low quality data that has been imputed in region (0-1)',
            'total_gaps_and_non_wear_percent_REGION': 'Total percentage of gaps and non-wear combined in region',
        }

        # Periphery quality features
        self._periphery_quality_features = {
            'total_duration_PERIPHERY': 'Total duration of the periphery (hours)',
            'device_turned_on_duration_PERIPHERY': 'Duration device was turned on in periphery (hours)',
            'device_turned_on_percent_PERIPHERY': 'Percentage of time device was turned on in periphery',
            'device_worn_duration_PERIPHERY': 'Duration device was worn in periphery (hours)',
            'device_worn_percent_PERIPHERY': 'Percentage of time device was worn in periphery',
            'imputed_duration_PERIPHERY': 'Duration of imputed data in periphery (hours)',
            'imputed_percent_PERIPHERY': 'Percentage of imputed data in periphery',
            'imputed_low_quality_duration_PERIPHERY': 'Duration of imputed low quality data in periphery (hours)',
            'imputed_low_quality_percent_PERIPHERY': 'Percentage of imputed low quality data in periphery',
            'unimputed_low_quality_duration_PERIPHERY': 'Duration of low quality data not imputed in periphery (hours)',
            'unimputed_low_quality_percent_PERIPHERY': 'Percentage of low quality data not imputed in periphery',
            'total_low_quality_duration_PERIPHERY': 'Total duration of low quality data in periphery (hours)',
            'total_low_quality_percent_PERIPHERY': 'Total percentage of low quality data in periphery',
            'imputed_gap_duration_PERIPHERY': 'Duration of imputed gap data in periphery (hours)',
            'imputed_gap_percent_PERIPHERY': 'Percentage of imputed gap data in periphery',
            'unimputed_gap_duration_PERIPHERY': 'Duration of unimputed gap data in periphery (hours)',
            'unimputed_gap_percent_PERIPHERY': 'Percentage of unimputed gap data in periphery',
            'total_gap_duration_PERIPHERY': 'Total duration of gap data in periphery (hours)',
            'total_gap_percent_PERIPHERY': 'Total percentage of gap data in periphery',
            'gap_imputation_ratio_PERIPHERY': 'Ratio of gap data that has been imputed in periphery (0-1)',
            'imputed_non_wear_duration_PERIPHERY': 'Duration of imputed non-wear data in periphery (hours)',
            'imputed_non_wear_percent_PERIPHERY': 'Percentage of imputed non-wear data in periphery',
            'unimputed_non_wear_duration_PERIPHERY': 'Duration of unimputed non-wear data in periphery (hours)',
            'unimputed_non_wear_percent_PERIPHERY': 'Percentage of unimputed non-wear data in periphery',
            'total_non_wear_duration_PERIPHERY': 'Total duration of non-wear data in periphery (hours)',
            'total_non_wear_percent_PERIPHERY': 'Total percentage of non-wear data in periphery',
            'non_wear_imputation_ratio_PERIPHERY': 'Ratio of non-wear data that has been imputed in periphery (0-1)',
            'imputed_jump_duration_PERIPHERY': 'Duration of imputed jump data in periphery (hours)',
            'imputed_jump_percent_PERIPHERY': 'Percentage of imputed jump data in periphery',
            'unimputed_jump_duration_PERIPHERY': 'Duration of unimputed jump data in periphery (hours)',
            'unimputed_jump_percent_PERIPHERY': 'Percentage of unimputed jump data in periphery',
            'total_jump_duration_PERIPHERY': 'Total duration of jump data in periphery (hours)',
            'total_jump_percent_PERIPHERY': 'Total percentage of jump data in periphery',
            'jump_imputation_ratio_PERIPHERY': 'Ratio of jump data that has been imputed in periphery (0-1)',
            'imputed_plummet_duration_PERIPHERY': 'Duration of imputed plummet data in periphery (hours)',
            'imputed_plummet_percent_PERIPHERY': 'Percentage of imputed plummet data in periphery',
            'unimputed_plummet_duration_PERIPHERY': 'Duration of unimputed plummet data in periphery (hours)',
            'unimputed_plummet_percent_PERIPHERY': 'Percentage of unimputed plummet data in periphery',
            'total_plummet_duration_PERIPHERY': 'Total duration of plummet data in periphery (hours)',
            'total_plummet_percent_PERIPHERY': 'Total percentage of plummet data in periphery',
            'plummet_imputation_ratio_PERIPHERY': 'Ratio of plummet data that has been imputed in periphery (0-1)',
            'imputed_extreme_negative_duration_PERIPHERY': 'Duration of imputed extreme negative data in periphery (hours)',
            'imputed_extreme_negative_percent_PERIPHERY': 'Percentage of imputed extreme negative data in periphery',
            'unimputed_extreme_negative_duration_PERIPHERY': 'Duration of unimputed extreme negative data in periphery (hours)',
            'unimputed_extreme_negative_percent_PERIPHERY': 'Percentage of unimputed extreme negative data in periphery',
            'total_extreme_negative_duration_PERIPHERY': 'Total duration of extreme negative data in periphery (hours)',
            'total_extreme_negative_percent_PERIPHERY': 'Total percentage of extreme negative data in periphery',
            'sub_negative_10_sum_PERIPHERY': 'Sum of all TAC values below -10 in periphery',
            'extreme_negative_imputation_ratio_PERIPHERY': 'Ratio of extreme negative data that has been imputed in periphery (0-1)',
            'low_quality_imputation_ratio_PERIPHERY': 'Ratio of overall low quality data that has been imputed in periphery (0-1)',
            'total_gaps_and_non_wear_percent_PERIPHERY': 'Total percentage of gaps and non-wear combined in periphery',
        }

        # Periphery Before quality features
        self._periphery_before_quality_features = {
            'total_duration_PERIPHERY_BEFORE': 'Total duration of the periphery before the curve (hours)',
            'device_turned_on_duration_PERIPHERY_BEFORE': 'Duration device was turned on in periphery before (hours)',
            'device_turned_on_percent_PERIPHERY_BEFORE': 'Percentage of time device was turned on in periphery before',
            'device_worn_duration_PERIPHERY_BEFORE': 'Duration device was worn in periphery before (hours)',
            'device_worn_percent_PERIPHERY_BEFORE': 'Percentage of time device was worn in periphery before',
            'imputed_duration_PERIPHERY_BEFORE': 'Duration of imputed data in periphery before (hours)',
            'imputed_percent_PERIPHERY_BEFORE': 'Percentage of imputed data in periphery before',
            'imputed_low_quality_duration_PERIPHERY_BEFORE': 'Duration of imputed low quality data in periphery before (hours)',
            'imputed_low_quality_percent_PERIPHERY_BEFORE': 'Percentage of imputed low quality data in periphery before',
            'unimputed_low_quality_duration_PERIPHERY_BEFORE': 'Duration of low quality data not imputed in periphery before (hours)',
            'unimputed_low_quality_percent_PERIPHERY_BEFORE': 'Percentage of low quality data not imputed in periphery before',
            'total_low_quality_duration_PERIPHERY_BEFORE': 'Total duration of low quality data in periphery before (hours)',
            'total_low_quality_percent_PERIPHERY_BEFORE': 'Total percentage of low quality data in periphery before',
            'imputed_gap_duration_PERIPHERY_BEFORE': 'Duration of imputed gap data in periphery before (hours)',
            'imputed_gap_percent_PERIPHERY_BEFORE': 'Percentage of imputed gap data in periphery before',
            'unimputed_gap_duration_PERIPHERY_BEFORE': 'Duration of unimputed gap data in periphery before (hours)',
            'unimputed_gap_percent_PERIPHERY_BEFORE': 'Percentage of unimputed gap data in periphery before',
            'total_gap_duration_PERIPHERY_BEFORE': 'Total duration of gap data in periphery before (hours)',
            'total_gap_percent_PERIPHERY_BEFORE': 'Total percentage of gap data in periphery before',
            'total_gap_duration_inclusive_PERIPHERY_BEFORE': 'Total duration of gap data in periphery before (hours); inclusive: gap==1 only, no mutual exclusivity',
            'total_gap_percent_inclusive_PERIPHERY_BEFORE': 'Total percentage of gap data in periphery before; inclusive: gap==1 only, no mutual exclusivity',
            'gap_imputation_ratio_PERIPHERY_BEFORE': 'Ratio of gap data that has been imputed in periphery before (0-1)',
            'imputed_non_wear_duration_PERIPHERY_BEFORE': 'Duration of imputed non-wear data in periphery before (hours)',
            'imputed_non_wear_percent_PERIPHERY_BEFORE': 'Percentage of imputed non-wear data in periphery before',
            'unimputed_non_wear_duration_PERIPHERY_BEFORE': 'Duration of unimputed non-wear data in periphery before (hours)',
            'unimputed_non_wear_percent_PERIPHERY_BEFORE': 'Percentage of unimputed non-wear data in periphery before',
            'total_non_wear_duration_PERIPHERY_BEFORE': 'Total duration of non-wear data in periphery before (hours)',
            'total_non_wear_percent_PERIPHERY_BEFORE': 'Total percentage of non-wear data in periphery before',
            'total_non_wear_duration_inclusive_PERIPHERY_BEFORE': 'Total duration of non-wear data in periphery before (hours); inclusive: non_wear==1 only, no mutual exclusivity',
            'total_non_wear_percent_inclusive_PERIPHERY_BEFORE': 'Total percentage of non-wear data in periphery before; inclusive: non_wear==1 only, no mutual exclusivity',
            'non_wear_imputation_ratio_PERIPHERY_BEFORE': 'Ratio of non-wear data that has been imputed in periphery before (0-1)',
            'imputed_jump_duration_PERIPHERY_BEFORE': 'Duration of imputed jump data in periphery before (hours)',
            'imputed_jump_percent_PERIPHERY_BEFORE': 'Percentage of imputed jump data in periphery before',
            'unimputed_jump_duration_PERIPHERY_BEFORE': 'Duration of unimputed jump data in periphery before (hours)',
            'unimputed_jump_percent_PERIPHERY_BEFORE': 'Percentage of unimputed jump data in periphery before',
            'total_jump_duration_PERIPHERY_BEFORE': 'Total duration of jump data in periphery before (hours)',
            'total_jump_percent_PERIPHERY_BEFORE': 'Total percentage of jump data in periphery before',
            'total_jump_duration_inclusive_PERIPHERY_BEFORE': 'Total duration of jump data in periphery before (hours); inclusive: jump==1 only, no mutual exclusivity',
            'total_jump_percent_inclusive_PERIPHERY_BEFORE': 'Total percentage of jump data in periphery before; inclusive: jump==1 only, no mutual exclusivity',
            'jump_imputation_ratio_PERIPHERY_BEFORE': 'Ratio of jump data that has been imputed in periphery before (0-1)',
            'imputed_plummet_duration_PERIPHERY_BEFORE': 'Duration of imputed plummet data in periphery before (hours)',
            'imputed_plummet_percent_PERIPHERY_BEFORE': 'Percentage of imputed plummet data in periphery before',
            'unimputed_plummet_duration_PERIPHERY_BEFORE': 'Duration of unimputed plummet data in periphery before (hours)',
            'unimputed_plummet_percent_PERIPHERY_BEFORE': 'Percentage of unimputed plummet data in periphery before',
            'total_plummet_duration_PERIPHERY_BEFORE': 'Total duration of plummet data in periphery before (hours)',
            'total_plummet_percent_PERIPHERY_BEFORE': 'Total percentage of plummet data in periphery before',
            'total_plummet_duration_inclusive_PERIPHERY_BEFORE': 'Total duration of plummet data in periphery before (hours); inclusive: plummet==1 only, no mutual exclusivity',
            'total_plummet_percent_inclusive_PERIPHERY_BEFORE': 'Total percentage of plummet data in periphery before; inclusive: plummet==1 only, no mutual exclusivity',
            'plummet_imputation_ratio_PERIPHERY_BEFORE': 'Ratio of plummet data that has been imputed in periphery before (0-1)',
            'imputed_extreme_negative_duration_PERIPHERY_BEFORE': 'Duration of imputed extreme negative data in periphery before (hours)',
            'imputed_extreme_negative_percent_PERIPHERY_BEFORE': 'Percentage of imputed extreme negative data in periphery before',
            'unimputed_extreme_negative_duration_PERIPHERY_BEFORE': 'Duration of unimputed extreme negative data in periphery before (hours)',
            'unimputed_extreme_negative_percent_PERIPHERY_BEFORE': 'Percentage of unimputed extreme negative data in periphery before',
            'total_extreme_negative_duration_PERIPHERY_BEFORE': 'Total duration of extreme negative data in periphery before (hours)',
            'total_extreme_negative_percent_PERIPHERY_BEFORE': 'Total percentage of extreme negative data in periphery before',
            'total_extreme_negative_duration_inclusive_PERIPHERY_BEFORE': 'Total duration of extreme negative data in periphery before (hours); inclusive: extreme_negative==1 only, no mutual exclusivity',
            'total_extreme_negative_percent_inclusive_PERIPHERY_BEFORE': 'Total percentage of extreme negative data in periphery before; inclusive: extreme_negative==1 only, no mutual exclusivity',
            'sub_negative_10_sum_PERIPHERY_BEFORE': 'Sum of all TAC values below -10 in periphery before',
            'extreme_negative_imputation_ratio_PERIPHERY_BEFORE': 'Ratio of extreme negative data that has been imputed in periphery before (0-1)',
            'low_quality_imputation_ratio_PERIPHERY_BEFORE': 'Ratio of overall low quality data that has been imputed in periphery before (0-1)',
            'total_gaps_and_non_wear_percent_PERIPHERY_BEFORE': 'Total percentage of gaps and non-wear combined in periphery before',
        }

        # Periphery After quality features  
        self._periphery_after_quality_features = {
            'total_duration_PERIPHERY_AFTER': 'Total duration of the periphery after the curve (hours)',
            'device_turned_on_duration_PERIPHERY_AFTER': 'Duration device was turned on in periphery after (hours)',
            'device_turned_on_percent_PERIPHERY_AFTER': 'Percentage of time device was turned on in periphery after',
            'device_worn_duration_PERIPHERY_AFTER': 'Duration device was worn in periphery after (hours)',
            'device_worn_percent_PERIPHERY_AFTER': 'Percentage of time device was worn in periphery after',
            'imputed_duration_PERIPHERY_AFTER': 'Duration of imputed data in periphery after (hours)',
            'imputed_percent_PERIPHERY_AFTER': 'Percentage of imputed data in periphery after',
            'imputed_low_quality_duration_PERIPHERY_AFTER': 'Duration of imputed low quality data in periphery after (hours)',
            'imputed_low_quality_percent_PERIPHERY_AFTER': 'Percentage of imputed low quality data in periphery after',
            'unimputed_low_quality_duration_PERIPHERY_AFTER': 'Duration of low quality data not imputed in periphery after (hours)',
            'unimputed_low_quality_percent_PERIPHERY_AFTER': 'Percentage of low quality data not imputed in periphery after',
            'total_low_quality_duration_PERIPHERY_AFTER': 'Total duration of low quality data in periphery after (hours)',
            'total_low_quality_percent_PERIPHERY_AFTER': 'Total percentage of low quality data in periphery after',
            'imputed_gap_duration_PERIPHERY_AFTER': 'Duration of imputed gap data in periphery after (hours)',
            'imputed_gap_percent_PERIPHERY_AFTER': 'Percentage of imputed gap data in periphery after',
            'unimputed_gap_duration_PERIPHERY_AFTER': 'Duration of unimputed gap data in periphery after (hours)',
            'unimputed_gap_percent_PERIPHERY_AFTER': 'Percentage of unimputed gap data in periphery after',
            'total_gap_duration_PERIPHERY_AFTER': 'Total duration of gap data in periphery after (hours)',
            'total_gap_percent_PERIPHERY_AFTER': 'Total percentage of gap data in periphery after',
            'total_gap_duration_inclusive_PERIPHERY_AFTER': 'Total duration of gap data in periphery after (hours); inclusive: gap==1 only, no mutual exclusivity',
            'total_gap_percent_inclusive_PERIPHERY_AFTER': 'Total percentage of gap data in periphery after; inclusive: gap==1 only, no mutual exclusivity',
            'gap_imputation_ratio_PERIPHERY_AFTER': 'Ratio of gap data that has been imputed in periphery after (0-1)',
            'imputed_non_wear_duration_PERIPHERY_AFTER': 'Duration of imputed non-wear data in periphery after (hours)',
            'imputed_non_wear_percent_PERIPHERY_AFTER': 'Percentage of imputed non-wear data in periphery after',
            'unimputed_non_wear_duration_PERIPHERY_AFTER': 'Duration of unimputed non-wear data in periphery after (hours)',
            'unimputed_non_wear_percent_PERIPHERY_AFTER': 'Percentage of unimputed non-wear data in periphery after',
            'total_non_wear_duration_PERIPHERY_AFTER': 'Total duration of non-wear data in periphery after (hours)',
            'total_non_wear_percent_PERIPHERY_AFTER': 'Total percentage of non-wear data in periphery after',
            'total_non_wear_duration_inclusive_PERIPHERY_AFTER': 'Total duration of non-wear data in periphery after (hours); inclusive: non_wear==1 only, no mutual exclusivity',
            'total_non_wear_percent_inclusive_PERIPHERY_AFTER': 'Total percentage of non-wear data in periphery after; inclusive: non_wear==1 only, no mutual exclusivity',
            'non_wear_imputation_ratio_PERIPHERY_AFTER': 'Ratio of non-wear data that has been imputed in periphery after (0-1)',
            'imputed_jump_duration_PERIPHERY_AFTER': 'Duration of imputed jump data in periphery after (hours)',
            'imputed_jump_percent_PERIPHERY_AFTER': 'Percentage of imputed jump data in periphery after',
            'unimputed_jump_duration_PERIPHERY_AFTER': 'Duration of unimputed jump data in periphery after (hours)',
            'unimputed_jump_percent_PERIPHERY_AFTER': 'Percentage of unimputed jump data in periphery after',
            'total_jump_duration_PERIPHERY_AFTER': 'Total duration of jump data in periphery after (hours)',
            'total_jump_percent_PERIPHERY_AFTER': 'Total percentage of jump data in periphery after',
            'total_jump_duration_inclusive_PERIPHERY_AFTER': 'Total duration of jump data in periphery after (hours); inclusive: jump==1 only, no mutual exclusivity',
            'total_jump_percent_inclusive_PERIPHERY_AFTER': 'Total percentage of jump data in periphery after; inclusive: jump==1 only, no mutual exclusivity',
            'jump_imputation_ratio_PERIPHERY_AFTER': 'Ratio of jump data that has been imputed in periphery after (0-1)',
            'imputed_plummet_duration_PERIPHERY_AFTER': 'Duration of imputed plummet data in periphery after (hours)',
            'imputed_plummet_percent_PERIPHERY_AFTER': 'Percentage of imputed plummet data in periphery after',
            'unimputed_plummet_duration_PERIPHERY_AFTER': 'Duration of unimputed plummet data in periphery after (hours)',
            'unimputed_plummet_percent_PERIPHERY_AFTER': 'Percentage of unimputed plummet data in periphery after',
            'total_plummet_duration_PERIPHERY_AFTER': 'Total duration of plummet data in periphery after (hours)',
            'total_plummet_percent_PERIPHERY_AFTER': 'Total percentage of plummet data in periphery after',
            'total_plummet_duration_inclusive_PERIPHERY_AFTER': 'Total duration of plummet data in periphery after (hours); inclusive: plummet==1 only, no mutual exclusivity',
            'total_plummet_percent_inclusive_PERIPHERY_AFTER': 'Total percentage of plummet data in periphery after; inclusive: plummet==1 only, no mutual exclusivity',
            'plummet_imputation_ratio_PERIPHERY_AFTER': 'Ratio of plummet data that has been imputed in periphery after (0-1)',
            'imputed_extreme_negative_duration_PERIPHERY_AFTER': 'Duration of imputed extreme negative data in periphery after (hours)',
            'imputed_extreme_negative_percent_PERIPHERY_AFTER': 'Percentage of imputed extreme negative data in periphery after',
            'unimputed_extreme_negative_duration_PERIPHERY_AFTER': 'Duration of unimputed extreme negative data in periphery after (hours)',
            'unimputed_extreme_negative_percent_PERIPHERY_AFTER': 'Percentage of unimputed extreme negative data in periphery after',
            'total_extreme_negative_duration_PERIPHERY_AFTER': 'Total duration of extreme negative data in periphery after (hours)',
            'total_extreme_negative_percent_PERIPHERY_AFTER': 'Total percentage of extreme negative data in periphery after',
            'total_extreme_negative_duration_inclusive_PERIPHERY_AFTER': 'Total duration of extreme negative data in periphery after (hours); inclusive: extreme_negative==1 only, no mutual exclusivity',
            'total_extreme_negative_percent_inclusive_PERIPHERY_AFTER': 'Total percentage of extreme negative data in periphery after; inclusive: extreme_negative==1 only, no mutual exclusivity',
            'sub_negative_10_sum_PERIPHERY_AFTER': 'Sum of all TAC values below -10 in periphery after',
            'extreme_negative_imputation_ratio_PERIPHERY_AFTER': 'Ratio of extreme negative data that has been imputed in periphery after (0-1)',
            'low_quality_imputation_ratio_PERIPHERY_AFTER': 'Ratio of overall low quality data that has been imputed in periphery after (0-1)',
            'total_gaps_and_non_wear_percent_PERIPHERY_AFTER': 'Total percentage of gaps and non-wear combined in periphery after',
        }

    @property
    def curve_feature_descriptions(self):
        """Get curve feature descriptions."""
        return {
            **self._curve_feature_descriptions,
            **self._region_quality_features,
            **self._periphery_quality_features,
            **self._periphery_before_quality_features,
            **self._periphery_after_quality_features
        }

    @property
    def curve_features_descriptions(self):
        """Get curve feature descriptions."""
        return self._curve_feature_descriptions

    @property
    def event_feature_descriptions(self):
        """Get event feature descriptions."""
        return OrderedDict(self._event_feature_descriptions)

    @property
    def person_level_feature_descriptions(self):
        """Get person-level feature descriptions."""
        return self._person_level_feature_descriptions

    @property
    def imputation_feature_descriptions(self):
        """Get imputation feature descriptions."""
        return self._imputation_feature_descriptions

    @property
    def day_feature_descriptions(self):
        """Get day feature descriptions."""
        return self._day_feature_descriptions

    @property
    def curve_features_with_events_descriptions(self):
        """Get all feature descriptions combined."""
        return {
            **self._curve_feature_descriptions,
            **self._person_level_feature_descriptions,
            **self._event_feature_descriptions,
            **self._imputation_feature_descriptions,
            **self._region_quality_features,
            **self._periphery_quality_features,
            **self._periphery_before_quality_features,
            **self._periphery_after_quality_features
        }

    @property
    def person_level_feature_types(self):
        """Get person-level feature types."""
        return self._person_level_feature_types

    @property
    def stats_features(self):
        """Get features used for statistical analysis."""
        return self._stats_features

    @property
    def region_definitions(self):
        """Get region definitions."""
        return self._region_definitions

    @property
    def region_quality_features(self):
        """Get region quality features."""
        return self._region_quality_features

    @property
    def periphery_quality_features(self):
        """Get periphery quality features."""
        return self._periphery_quality_features

    @property
    def periphery_before_quality_features(self):
        """Get periphery before quality features."""
        return self._periphery_before_quality_features

    @property
    def periphery_after_quality_features(self):
        """Get periphery after quality features."""
        return self._periphery_after_quality_features

    @property
    def periphery_before_feature_types(self):
        """Get periphery before feature types."""
        return self._periphery_before_feature_types

    @property
    def periphery_after_feature_types(self):
        """Get periphery after feature types."""
        return self._periphery_after_feature_types

    @classmethod
    def get_variable_key_dataframes(cls):
        """Get a list of (name, dataframe) tuples for variable keys in export order."""
        instance = cls()
        dfs = []
        
        # Region definitions
        region_df = pd.DataFrame({
            'Variable': list(instance.region_definitions.keys()),
            'Description': list(instance.region_definitions.values())
        })
        region_df.set_index('Variable', inplace=True)
        region_df.index.name = None
        region_df.columns = pd.MultiIndex.from_product([['Region Definitions'], region_df.columns])
        dfs.append(('Region Definitions', region_df))
        
        # Features tab
        features_df = pd.DataFrame({
            'Variable': list(instance.curve_feature_descriptions.keys()),
            'Description': list(instance.curve_feature_descriptions.values())
        })
        features_df.set_index('Variable', inplace=True)
        features_df.index.name = None
        features_df.columns = pd.MultiIndex.from_product([['Features'], features_df.columns])
        dfs.append(('Features', features_df))
        
        # Events tab
        events_df = pd.DataFrame({
            'Variable': list(instance.event_feature_descriptions.keys()),
            'Description': list(instance.event_feature_descriptions.values())
        })
        events_df.set_index('Variable', inplace=True)
        events_df.index.name = None
        events_df.columns = pd.MultiIndex.from_product([['Events'], events_df.columns])
        dfs.append(('Events', events_df))
        
        # Valid Matched by SubID tab
        person_df = pd.DataFrame({
            'Variable': list(instance.person_level_feature_descriptions.keys()),
            'Description': list(instance.person_level_feature_descriptions.values())
        })
        person_df.set_index('Variable', inplace=True)
        person_df.index.name = None
        person_df.columns = pd.MultiIndex.from_product([['Valid Matched by SubID'], person_df.columns])
        dfs.append(('Valid Matched by SubID', person_df))
        
        # Imputations tab
        imputations_df = pd.DataFrame({
            'Variable': list(instance.imputation_feature_descriptions.keys()),
            'Description': list(instance.imputation_feature_descriptions.values())
        })
        imputations_df.set_index('Variable', inplace=True)
        imputations_df.index.name = None
        imputations_df.columns = pd.MultiIndex.from_product([['Imputations'], imputations_df.columns])
        dfs.append(('Imputations', imputations_df))
        
        # Day Features tab
        day_df = pd.DataFrame({
            'Variable': list(instance.day_feature_descriptions.keys()),
            'Description': list(instance.day_feature_descriptions.values())
        })
        day_df.set_index('Variable', inplace=True)
        day_df.index.name = None
        day_df.columns = pd.MultiIndex.from_product([['Day Features'], day_df.columns])
        dfs.append(('Day Features', day_df))
        
        return dfs

    @classmethod
    def get_run_settings_dataframe(cls, smooth_and_impute_attrs=None, curve_attrs=None, raw_curve_attrs=None, event_attrs=None, day_attrs=None):
        """Get a DataFrame of run settings in the correct format for export."""
        run_settings = []
        
        # Helper function to add settings
        def add_setting(category, setting, value):
            run_settings.append({
                'Category': category,
                'Setting': setting,
                'Value': str(value)  # Convert all values to strings for consistent display
            })

        # Add Smooth and Impute settings
        if smooth_and_impute_attrs:
            for key, value in smooth_and_impute_attrs.items():
                add_setting('Smooth and Impute', key, value)
        
        # Add Curve settings
        if curve_attrs:
            for key, value in curve_attrs.items():
                if isinstance(value, dict):
                    for subkey, subvalue in value.items():
                        if isinstance(subvalue, dict):
                            for subsubkey, subsubvalue in subvalue.items():
                                add_setting(f'Curve - {key}', f"{subkey}.{subsubkey}", subsubvalue)
                        else:
                            add_setting(f'Curve - {key}', subkey, subvalue)
                else:
                    add_setting('Curve', key, value)

        # Add Raw Curve settings
        if raw_curve_attrs:
            for key, value in raw_curve_attrs.items():
                add_setting('Raw Curve', key, value)

        # Add Event settings
        if event_attrs:
            for key, value in event_attrs.items():
                if isinstance(value, list):
                    add_setting('Event', key, ', '.join(str(v) for v in value))
                else:
                    add_setting('Event', key, value)

        # Add Day settings
        if day_attrs:
            for key, value in day_attrs.items():
                add_setting('Day', key, value)

        return pd.DataFrame(run_settings) if run_settings else None

    def get_tab_descriptions_dataframe(self, include_events=False):
        """
        Get a DataFrame describing the tabs in the exported workbook.
        
        Args:
            include_events (bool): Whether to include event-related tabs
            
        Returns:
            pd.DataFrame: DataFrame with tab names and descriptions
        """
        descriptions = self._curve_features_with_events_tab_descriptions if include_events else self._curve_features_tab_descriptions
        return pd.DataFrame({
            'Tab Name': list(descriptions.keys()),
            'Description': list(descriptions.values())
        })

# Create a singleton instance
report_guide = ReportGuide()

# Export the properties for backward compatibility
curve_feature_descriptions = report_guide.curve_feature_descriptions
event_feature_descriptions = report_guide.event_feature_descriptions
person_level_feature_descriptions = report_guide.person_level_feature_descriptions
imputation_feature_descriptions = report_guide.imputation_feature_descriptions
day_feature_descriptions = report_guide.day_feature_descriptions
curve_features_descriptions = report_guide.curve_features_descriptions
curve_features_with_events_descriptions = report_guide.curve_features_with_events_descriptions
person_level_feature_types = report_guide.person_level_feature_types
stats_features = report_guide.stats_features 