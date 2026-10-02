from __future__ import annotations

import warnings

from App.SDM.Analysis.statModel import statModel
from App.SDM.Configuration.file_management import load, save_to_computer
from App.SDM.Documenting.embed_graphs import embed_graphs_into_workbook_tab
from App.SDM.Configuration.file_management import extract_subid
import os
import re
from typing import List, Optional, Tuple
import pandas as pd
import numpy as np
from openpyxl import Workbook
from openpyxl.drawing.image import Image

# Must match ``dayFeatures.TOP_HQ_CURVE_COL_PREFIX`` (used where class attrs are not visible to static checks).
_TOP_HQ_CURVE_COL_PREFIX = 'top_hq_curve_'

# Max-HQ overlap selection: drop prior-day-starting curves with little of the curve on the social day.
_HQ_PICK_PRIOR_DAY_MIN_OVERLAP_HOURS = 1.0
_HQ_PICK_PRIOR_DAY_MIN_CURVE_FRACTION_IN_DAY = 0.5

# Prior-day medians: only count a prior day if its max-HQ curve has at least this much HQ time (hours).
_PRIOR_MEDIAN_MIN_HIGH_QUALITY_DURATION_HOURS = 1.0

# ---------------------------------------------------------------------------
# Per-slot curve feature attachment (replaces single top_hq merge)
# ---------------------------------------------------------------------------
_MAX_CURVE_SLOTS = 5

# Curve feature suffixes to copy per overlap slot (beyond the 7 overlap-metadata fields).
_PER_SLOT_CURVE_FEATURE_SUFFIXES: tuple[str, ...] = (
    'peak_CURVE',
    'auc_total_CURVE',
    'auc_relative_CURVE',
    'curve_threshold',
    'high_quality_percent_CURVE',
    'total_duration_CURVE',
    'rise_rate_CURVE',
    'fall_rate_CURVE',
    'rise_rate_point_to_point_CURVE',
    'fall_rate_point_to_point_CURVE',
    'rise_duration_CURVE',
    'fall_duration_CURVE',
    'below_threshold_percent_CURVE',
    'started_curve_count_CURVE',
    # Quality (curve segment)
    'device_turned_on_percent_CURVE',
    'device_worn_percent_CURVE',
    'imputed_percent_CURVE',
    'total_low_quality_percent_CURVE',
    'unimputed_low_quality_percent_CURVE',
    'total_gap_percent_CURVE',
    'total_non_wear_percent_CURVE',
    'total_jump_percent_CURVE',
    'total_plummet_percent_CURVE',
    'total_extreme_negative_percent_CURVE',
    'low_quality_imputation_ratio_CURVE',
    # Periphery before
    'total_duration_PERIPHERY_BEFORE',
    'device_turned_on_percent_PERIPHERY_BEFORE',
    'device_worn_percent_PERIPHERY_BEFORE',
    'imputed_percent_PERIPHERY_BEFORE',
    'total_low_quality_percent_PERIPHERY_BEFORE',
    'unimputed_low_quality_percent_PERIPHERY_BEFORE',
    'total_gap_percent_PERIPHERY_BEFORE',
    'total_non_wear_percent_PERIPHERY_BEFORE',
    'total_jump_percent_PERIPHERY_BEFORE',
    'total_plummet_percent_PERIPHERY_BEFORE',
    'total_extreme_negative_percent_PERIPHERY_BEFORE',
    'low_quality_imputation_ratio_PERIPHERY_BEFORE',
    # Periphery after
    'total_duration_PERIPHERY_AFTER',
    'device_turned_on_percent_PERIPHERY_AFTER',
    'device_worn_percent_PERIPHERY_AFTER',
    'imputed_percent_PERIPHERY_AFTER',
    'total_low_quality_percent_PERIPHERY_AFTER',
    'unimputed_low_quality_percent_PERIPHERY_AFTER',
    'total_gap_percent_PERIPHERY_AFTER',
    'total_non_wear_percent_PERIPHERY_AFTER',
    'total_jump_percent_PERIPHERY_AFTER',
    'total_plummet_percent_PERIPHERY_AFTER',
    'total_extreme_negative_percent_PERIPHERY_AFTER',
    'low_quality_imputation_ratio_PERIPHERY_AFTER',
)

# Median-prior: qualifying curve thresholds
_MEDIAN_PRIOR_MIN_HQ_PERCENT = 0.75
_MEDIAN_PRIOR_MIN_HQ_DURATION_HOURS = 1.0

_PRIOR_MEDIAN_CURVE_METRICS_V2 = (
    'peak_CURVE',
    'auc_total_CURVE',
    'duration_CURVE',
    'total_low_quality_percent_CURVE',
    'rise_rate_CURVE',
    'fall_rate_CURVE',
    'rise_rate_point_to_point_CURVE',
    'fall_rate_point_to_point_CURVE',
    'fall_duration_CURVE',
    'rise_duration_CURVE',
)


class dayFeatures():
    # ``self_report_and_tac_comparison`` value keys with burst-tab labels (also used in Agreement_summary grid cells).
    HQ_AGREEMENT_ORDER_AND_LABELS = (
        ('true_positive', 'TP'),
        ('true_negative', 'TN'),
        ('false_positive', 'FP'),
        ('false_negative', 'FN'),
        ('true_unknown', 'True unknown'),
        ('false_unknown', 'False unknown'),
        ('unknown_sr_missing', 'unknown [SR missing]'),
        ('positive_without_self_report', 'Positive TAC, no morning SR'),
        ('negative_without_self_report', 'Negative TAC, no morning SR'),
    )

    # Curve-level fields merged onto each day from the overlap curve with max HQ duration (distinct from
    # time-ordered overlap slots ``curve_1_id``, ``curve_2_id``, …).
    TOP_HQ_CURVE_COL_PREFIX = _TOP_HQ_CURVE_COL_PREFIX

    def __init__(self, processed_data_folder, subids=None):
        print(f"Looking for files in: {processed_data_folder}")
        files = [f for f in os.listdir(processed_data_folder) 
                 if 'processed' in f and not f.startswith('.')]
        
        # Filter files by subid if specified
        if subids is not None:
            subids_str = set(str(s) for s in subids)
            files = [file for file in files if extract_subid(file) in subids_str]
            if not files:
                raise ValueError(f"No processed files found for subids {subids}")
            print(f"Filtered to files for subids {subids}: {files}")
        else:
            print(f"Found files: {files}")
        
        self.processors = [load(file[:-4], processed_data_folder) for file in files]
        print(f"Found {len(self.processors)} processor files")
        
        self.processors = [processor for processor in self.processors if hasattr(processor, 'day_level_data')]
        print(f"Found {len(self.processors)} processors with day_level_data")
        
        if len(self.processors) == 0:
            raise ValueError("No processors with day_level_data found")
            
        # Print summary about processors
        print(f"Loaded {len(self.processors)} processors with day-level data")
            
        self.day_features = pd.concat([processor.day_level_data for processor in self.processors])
        
        # Drop rows with missing SubID, Dataset_ID, or day_no before converting to int
        rows_before = len(self.day_features)
        self.day_features.dropna(subset=['SubID', 'Dataset_ID', 'day_no'], inplace=True)
        self.day_features[['SubID', 'day_no']] = self.day_features[['SubID', 'day_no']].astype(int)
        # Primary key: SubID + Dataset_ID + day_no
        self.day_features.drop_duplicates(subset=['SubID', 'Dataset_ID', 'day_no'], inplace=True)
        print(f"Day features: {len(self.day_features)} rows ({rows_before - len(self.day_features)} dropped)")
        self.day_stat_frames = []

        # Early: weekday code from social-day start (ISO-style 1=Monday … 7=Sunday)
        dayFeatures.assign_day_of_week_cat_from_begin_day(self.day_features)
        dayFeatures.assign_new_device_flag(self.day_features)

        # Split data into valid and invalid days if there's a validity column
        if 'DAY_VALID' in self.day_features.columns:
            self.day_valid = self.day_features[self.day_features['DAY_VALID'] == 1]
            self.day_invalid = self.day_features[self.day_features['DAY_VALID'] != 1]
        else:
            self.day_valid = self.day_features
            self.day_invalid = pd.DataFrame()

        self.morning_annotated = None
        # One row per (SubID, Dataset_ID): skyn burst windows + optional global study-day bounds (set in filter_days_by_date_range).
        self.skyn_dates_metadata = None
        # Built by prepare_morning_self_report_for_tac_merge; consumed by add_morning_report_drink_agreement.
        self.morning_merge_key_calendar = None
        self.morning_merge_key_day_no = None

    @staticmethod
    def assign_day_of_week_cat_from_begin_day(df: pd.DataFrame) -> None:
        """
        Set ``day_of_week_cat`` from ``begin_day`` (mutates ``df`` in place).

        Uses the calendar date/time of ``begin_day``: **1 = Monday** through **7 = Sunday**
        (pandas ``Series.dt.dayofweek`` + 1). Nullable integer when ``begin_day`` is missing.
        """
        if df is None or len(df) == 0 or 'begin_day' not in df.columns:
            return
        df['begin_day'] = pd.to_datetime(df['begin_day'], errors='coerce')
        dow = df['begin_day'].dt.dayofweek
        df['day_of_week_cat'] = (dow + 1).astype(pd.Int64Dtype())

    @staticmethod
    def assign_new_device_flag(df: pd.DataFrame) -> None:
        """
        Set ``new_device`` = 1 when ``device_two`` is non-null, indicating that
        a device swap occurred during that day (two distinct devices contributed data).
        """
        if df is None or len(df) == 0:
            return
        if 'device_two' not in df.columns:
            df['new_device'] = 0
            return
        df['new_device'] = df['device_two'].notna().astype(int)

    def filter_days_by_date_range(self, metadata_csv_path, first_date_column='Day 1', last_date_column='Day 14', id_column='ID', burst_column='BURST', first_date_adjuster=0):
        """
        Filter day-level data to only include days within the date range specified in a metadata CSV.
        
        This method reads a CSV file with ID, burst, and date columns. It uses id_column as SubID and 
        burst_column as Dataset_ID to filter days. Only days with timestamps between first_date_column 
        and last_date_column (inclusive) are retained.
        
        Args:
            metadata_csv_path (str): Path to the metadata CSV file
            first_date_column (str): Column name for the start date (default: 'Day 1')
            last_date_column (str): Column name for the end date (default: 'Day 14')
            id_column (str): Column name for participant ID (default: 'ID')
            burst_column (str): Column name for burst/dataset ID (default: 'BURST')
            first_date_adjuster (int): Number of days to add to the first_date_column
                before filtering. For example, -1 will include the day before the
                recorded first_day (useful when days span 6am–6am and metadata
                marks the first full calendar day).
        
        Returns:
            None (modifies self.day_features in place)
        """
        print(f"\nFiltering days by metadata date range...")
        print(f"  Metadata file: {metadata_csv_path}")
        print(f"  Date range: {first_date_column} to {last_date_column}")
        
        _, ext = os.path.splitext(metadata_csv_path)
        if ext.lower() in ['.xlsx', '.xls']:
            metadata = pd.read_excel(metadata_csv_path)
        else:
            metadata = pd.read_csv(metadata_csv_path)
        
        # Strip whitespace from column names
        metadata.columns = metadata.columns.str.strip()

        # Check if required columns exist in metadata
        required_meta_cols = [id_column, burst_column, first_date_column, last_date_column]
        missing_meta_cols = [col for col in required_meta_cols if col not in metadata.columns]
        if missing_meta_cols:
            print(f"Warning: Missing required columns in metadata: {missing_meta_cols}")
            return
        
        # Forward-fill the ID column if there are null IDs (ID is only on the first row for each subject)
        if metadata[id_column].isna().any():
            print("  Detected null IDs in metadata - applying forward-fill...")
            metadata[id_column] = metadata[id_column].ffill()
        
        # Remove rows where ID or burst is still missing
        metadata = metadata.dropna(subset=[id_column, burst_column])
        
        # Convert date columns to datetime
        metadata[first_date_column] = pd.to_datetime(metadata[first_date_column], errors='coerce')
        metadata[last_date_column] = pd.to_datetime(metadata[last_date_column], errors='coerce')
        
        # Remove rows with invalid dates
        metadata = metadata.dropna(subset=[first_date_column, last_date_column])
        
        # Convert begin_day and end_day to datetime if not already
        self.day_features['begin_day'] = pd.to_datetime(self.day_features['begin_day'])
        self.day_features['end_day'] = pd.to_datetime(self.day_features['end_day'])
        
        # Reset index to ensure sequential integer indices
        self.day_features = self.day_features.reset_index(drop=True)
        
        # Create a list to track which rows to keep
        rows_to_keep = []
        rows_before = len(self.day_features)
        
        # Debug counters
        total_processed = 0
        matches_found = 0
        date_overlap_passed = 0
        
        # Process each day in day_features
        for idx, day_row in self.day_features.iterrows():
            total_processed += 1
            subid = str(day_row['SubID'])
            dataset_id = str(day_row['Dataset_ID'])
            day_start = day_row['begin_day']
            day_end = day_row['end_day']
            
            # Find matching metadata row
            # Convert Dataset_ID and BURST to int for comparison (Dataset_ID='001' should match BURST=1)
            try:
                dataset_id_int = int(dataset_id)
            except ValueError:
                # If Dataset_ID can't be converted to int, skip this row
                continue
            
            # Convert SubID to int for robust matching (handles leading zeros)
            try:
                subid_int = int(subid)
            except ValueError:
                # If SubID can't be converted to int, skip this row
                continue
            
            # Match on integer values to avoid string comparison issues
            metadata_row = metadata[
                (metadata[id_column].astype(int) == subid_int) &
                (metadata[burst_column].astype(int) == dataset_id_int)
            ]
            
            if metadata_row.empty:
                # No matching metadata found - exclude this day
                continue
            
            matches_found += 1
            
            # Get the date range for this SubID-Dataset_ID combination
            start_date = metadata_row[first_date_column].iloc[0]
            if first_date_adjuster != 0:
                start_date = start_date + pd.to_timedelta(first_date_adjuster, unit='D')
            end_date = metadata_row[last_date_column].iloc[0]
            
            # The metadata dates represent the START of each included day
            # CSV "Day 1" = 8/8/24 means include days starting on or after 8/8/24
            # CSV "Day 14" = 8/21/24 means include days starting on or before 8/21/24
            # So day_no=2 (begin_day=8/8) through day_no=15 (begin_day=8/21) will be kept
            if day_start >= start_date and day_start <= end_date:
                date_overlap_passed += 1
                rows_to_keep.append(idx)
        
        # Debug output
        print(f"  DEBUG - Total rows processed: {total_processed}")
        print(f"  DEBUG - Metadata matches found: {matches_found}")
        print(f"  DEBUG - Passed date overlap check: {date_overlap_passed}")
        print(f"  DEBUG - Items in rows_to_keep before dedup: {len(rows_to_keep)}")
        
        # Filter dataframe to only keep matching rows
        # Ensure unique indices (in case of duplicates)
        duplicates_count = len(rows_to_keep) - len(set(rows_to_keep))
        rows_to_keep = list(set(rows_to_keep))
        print(f"  DEBUG - Duplicate indices removed: {duplicates_count}")
        print(f"  DEBUG - Unique indices to keep: {len(rows_to_keep)}")
        
        if len(rows_to_keep) > 0:
            self.day_features = self.day_features.loc[rows_to_keep]
            self.day_features = self.day_features.reset_index(drop=True)
        else:
            # No rows to keep - empty the dataframe
            self.day_features = self.day_features.iloc[0:0]
        
        rows_after = len(self.day_features)
        rows_excluded = rows_before - rows_after
        
        print(f"  Days before filtering: {rows_before}")
        print(f"  Days after filtering: {rows_after}")
        print(f"  Days excluded: {rows_excluded}")
        
        # Update valid/invalid day splits if they exist
        if 'DAY_VALID' in self.day_features.columns:
            self.day_valid = self.day_features[self.day_features['DAY_VALID'] == 1]
            self.day_invalid = self.day_features[self.day_features['DAY_VALID'] != 1]
        else:
            self.day_valid = self.day_features
            self.day_invalid = pd.DataFrame()

        # Persist skyn-date metadata on each day row (same CSV used for filtering) for labeling / QC.
        meta_attach_cols = [id_column, burst_column, first_date_column, last_date_column]
        for _opt in ('study_day_first_day', 'study_day_last_day'):
            if _opt in metadata.columns and _opt not in meta_attach_cols:
                meta_attach_cols.append(_opt)
        meta_attach = metadata[meta_attach_cols].drop_duplicates().copy()
        meta_rename = {
            id_column: 'SubID',
            burst_column: 'Dataset_ID',
            first_date_column: 'metadata_skyn_first_day',
            last_date_column: 'metadata_skyn_last_day',
        }
        if 'study_day_first_day' in meta_attach.columns:
            meta_rename['study_day_first_day'] = 'metadata_study_day_first'
        if 'study_day_last_day' in meta_attach.columns:
            meta_rename['study_day_last_day'] = 'metadata_study_day_last'
        meta_attach = meta_attach.rename(columns=meta_rename)
        meta_attach['SubID'] = pd.to_numeric(meta_attach['SubID'], errors='coerce').astype(int)
        meta_attach['Dataset_ID'] = pd.to_numeric(meta_attach['Dataset_ID'], errors='coerce').astype(int)
        for _c in ('metadata_skyn_first_day', 'metadata_skyn_last_day'):
            meta_attach[_c] = pd.to_datetime(meta_attach[_c], errors='coerce')
        for _c in ('metadata_study_day_first', 'metadata_study_day_last'):
            if _c in meta_attach.columns:
                meta_attach[_c] = pd.to_numeric(meta_attach[_c], errors='coerce')

        self.skyn_dates_metadata = meta_attach.copy()

        if len(self.day_features) > 0:
            _drop_m = [
                'metadata_skyn_first_day',
                'metadata_skyn_last_day',
                'metadata_study_day_first',
                'metadata_study_day_last',
            ]
            for _c in _drop_m:
                if _c in self.day_features.columns:
                    self.day_features.drop(columns=[_c], inplace=True)
            self.day_features['SubID'] = pd.to_numeric(self.day_features['SubID'], errors='coerce').astype(int)
            self.day_features['Dataset_ID'] = pd.to_numeric(
                self.day_features['Dataset_ID'], errors='coerce'
            ).astype(int)
            self.day_features = self.day_features.merge(
                meta_attach, on=['SubID', 'Dataset_ID'], how='left'
            )
            print(
                "  Attached skyn metadata columns to day_features: "
                f"{', '.join(c for c in meta_rename.values() if c in self.day_features.columns)}"
            )

            # TAC-only: begin_day calendar vs skyn burst window (same inclusive rule as filtering)
            if (
                'metadata_skyn_first_day' in self.day_features.columns
                and 'begin_day' in self.day_features.columns
            ):
                _bd = pd.to_datetime(self.day_features['begin_day'], errors='coerce').dt.normalize()
                _w0 = (
                    pd.to_datetime(self.day_features['metadata_skyn_first_day'], errors='coerce')
                    + pd.to_timedelta(int(first_date_adjuster), unit='D')
                )
                _w0 = pd.to_datetime(_w0, errors='coerce').dt.normalize()
                _w1 = pd.to_datetime(
                    self.day_features['metadata_skyn_last_day'], errors='coerce'
                ).dt.normalize()
                _has_meta = self.day_features['metadata_skyn_first_day'].notna()
                _inside = _has_meta & _bd.notna() & (_bd >= _w0) & (_bd <= _w1)
                _outside = _has_meta & _bd.notna() & ~_inside
                if 'inside_burst' in self.day_features.columns:
                    self.day_features.drop(columns=['inside_burst'], inplace=True)
                self.day_features['inside_burst'] = np.nan
                self.day_features.loc[_inside, 'inside_burst'] = 1
                self.day_features.loc[_outside, 'inside_burst'] = 0
                print(
                    "  inside_burst (TAC begin_day vs skyn first/last): "
                    f"in=1: {int(_inside.sum())}, out=0: {int(_outside.sum())}, "
                    f"undetermined: {int(self.day_features['inside_burst'].isna().sum())}"
                )
    
    def filter_days_by_emadayn(self, min_day=1, max_day=28, device_filter=None):
        """
        Filter day-level data to only include study days within the specified EMADAYN range.
        This is useful for ACE and other studies that use EMADAYN to identify study days.
        
        Args:
            min_day (int): Minimum EMADAYN value to include (default: 1)
            max_day (int): Maximum EMADAYN value to include (default: 28)
            device_filter (str, optional): If provided, filter by device ID prefix (e.g., '31' or '32' for ACE)
        
        Returns:
            None (modifies self.day_features in place)
        """
        print(f"\nFiltering days by EMADAYN...")
        print(f"  EMADAYN range: {min_day} to {max_day}")
        
        rows_before = len(self.day_features)
        
        # Check if EMADAYN column exists
        if 'EMADAYN' not in self.day_features.columns:
            print(f"  Warning: EMADAYN column not found in day_features. Available columns: {list(self.day_features.columns)[:10]}...")
            print(f"  Skipping EMADAYN filter.")
            return
        
        # Filter by EMADAYN range
        self.day_features = self.day_features[
            (self.day_features['EMADAYN'] >= min_day) & 
            (self.day_features['EMADAYN'] <= max_day)
        ]
        
        rows_after_emadayn = len(self.day_features)
        print(f"  Days after EMADAYN filter: {rows_after_emadayn} (removed {rows_before - rows_after_emadayn})")
        
        # Optional device filter (for ACE: filter to newer devices with 31/32 prefix)
        if device_filter is not None:
            if 'device_one' in self.day_features.columns:
                device_col = 'device_one'
            elif 'device_id' in self.day_features.columns:
                device_col = 'device_id'
            else:
                print(f"  Warning: No device column found. Available columns: {list(self.day_features.columns)[:10]}...")
                print(f"  Skipping device filter.")
                device_col = None
            
            if device_col is not None:
                rows_before_device = len(self.day_features)
                # Convert device IDs to string and filter by prefix
                self.day_features[device_col] = self.day_features[device_col].astype(str)
                if isinstance(device_filter, list):
                    # Multiple prefixes (e.g., ['31', '32'])
                    mask = self.day_features[device_col].str.startswith(device_filter[0])
                    for prefix in device_filter[1:]:
                        mask = mask | self.day_features[device_col].str.startswith(prefix)
                    self.day_features = self.day_features[mask]
                else:
                    # Single prefix
                    self.day_features = self.day_features[
                        self.day_features[device_col].str.startswith(str(device_filter))
                    ]
                rows_after_device = len(self.day_features)
                print(f"  Days after device filter (prefix '{device_filter}'): {rows_after_device} (removed {rows_before_device - rows_after_device})")
        
        rows_after = len(self.day_features)
        rows_excluded = rows_before - rows_after
        
        print(f"  Days before filtering: {rows_before}")
        print(f"  Days after filtering: {rows_after}")
        print(f"  Days excluded: {rows_excluded}")
        
        # Update valid/invalid day splits if they exist
        if 'DAY_VALID' in self.day_features.columns:
            self.day_valid = self.day_features[self.day_features['DAY_VALID'] == 1]
            self.day_invalid = self.day_features[self.day_features['DAY_VALID'] != 1]
        else:
            self.day_valid = self.day_features
            self.day_invalid = pd.DataFrame()

    def compute_low_quality_stats(self, output_filepath=None):
        """
        Calculates the percentage of time in various low-quality states and the percentage of that time that was imputed.
        The results are stored in a dataframe and added to self.day_stat_frames.
        If an output_filepath is provided, the stats are also saved to a simple Excel file.
        """
        stats = {}
        
        # Define the low-quality categories and their corresponding column prefixes
        categories = {
            'Gaps': 'gap',
            'Non-Wear': 'non_wear',
            'Jumps': 'jump',
            'Plummets': 'plummet',
            'Extreme Negative': 'extreme_negative'
        }
        
        # Check if necessary columns exist
        required_cols = ['begin_day', 'end_day']
        for prefix in categories.values():
            required_cols.append(f'{prefix}_duration')
            required_cols.append(f'imputed_{prefix}_duration')

        missing_cols = [col for col in required_cols if col not in self.day_features.columns]
            
        if missing_cols:
            print(f"Warning: Missing required columns for low quality stats, skipping: {', '.join(missing_cols)}")
            self.day_stat_frames.append(pd.DataFrame()) # Append empty frame
            return

        # Calculate total day duration using time difference between begin_day and end_day
        # Convert to datetime if not already
        self.day_features['begin_day'] = pd.to_datetime(self.day_features['begin_day'])
        self.day_features['end_day'] = pd.to_datetime(self.day_features['end_day'])
        
        # Calculate duration in hours for each day
        day_durations = (self.day_features['end_day'] - self.day_features['begin_day']).dt.total_seconds() / 3600
        total_day_duration = day_durations.sum()

        for name, prefix in categories.items():
            duration_col = f'{prefix}_duration'
            imputed_duration_col = f'imputed_{prefix}_duration'

            total_category_duration = self.day_features[duration_col].sum()
            total_imputed_duration = self.day_features[imputed_duration_col].sum()
            
            percent_of_total_time = (total_category_duration / total_day_duration) * 100 if total_day_duration > 0 else 0
            
            percent_imputed = (total_imputed_duration / total_category_duration) * 100 if total_category_duration > 0 else 0
            
            stats[name] = {
                '% of Total Day Time': percent_of_total_time,
                '% Imputed': percent_imputed
            }
            
        stats_df = pd.DataFrame.from_dict(stats, orient='index')
        self.day_stat_frames.append(stats_df)

        if output_filepath:
            # Ensure the directory exists
            output_dir = os.path.dirname(output_filepath)
            if output_dir:
                os.makedirs(output_dir, exist_ok=True)
            stats_df.to_excel(output_filepath)

    def add_curve_overlap_detection(self, *args, **kwargs):
        """Deprecated. Use ``attach_curves_and_predict_drinking_days``."""
        warnings.warn(
            'add_curve_overlap_detection is deprecated. '
            'Use attach_curves_and_predict_drinking_days.',
            DeprecationWarning,
            stacklevel=2,
        )
        return self.attach_curves_and_predict_drinking_days(*args, **kwargs)

    def attach_curves_and_predict_drinking_days(
        self,
        curve_features_df=None,
        flag_non_wear_hours=6.0,
        flag_very_negative_hours=1.0,
        flag_artifact_hours=1.0,
        flag_proba_low=0.05,
        flag_proba_high=0.95,
    ):
        """
        Attach overlapping curves to each day and score the drinking-day model.

        Always creates ``_MAX_CURVE_SLOTS`` (5) overlap slots, each carrying the full set of
        per-slot curve features (``curve_{n}_<suffix>`` for every suffix in
        ``_PER_SLOT_CURVE_FEATURE_SUFFIXES``).  Slots are assigned by **descending
        high_quality_duration_CURVE** (slot 1 = most HQ data; tie-break: higher
        ``auc_total_CURVE``, then lower ``curve_id``); curves beyond slot 5 are ignored.

        Per slot, also computes:
        - ``curve_{n}_overlap_fraction`` = overlap_hours / duration_CURVE
        - ``curve_{n}_start_offset_hours`` = (begin_CURVE − begin_day) in hours (negative if prior day)
        - ``curve_{n}_end_offset_hours`` = (end_CURVE − begin_day) in hours (>24 if extends next day)

        After populating slots, computes ``median_prior_*`` across all qualifying prior
        curves in the burst (see ``assign_median_prior_from_all_qualifying_curves``).

        After curve slots are attached, scores the 08.05 drinking-day model
        (``pred_drinking_day``, ``proba_drinking_day``) and reliability flags. Threshold
        arguments are forwarded to that scorer. This is the day-level drinking label.

        Args:
            curve_features_df: DataFrame containing curve features (optional)
            flag_non_wear_hours: Flag when social-day hours not worn are at least this value.
                Hours not worn are ``day_hours - device_worn_duration`` (powered-off or missing
                time plus on-device non-wear).
            flag_very_negative_hours: Flag when ``extreme_negative_duration`` (TAC < -15) is at least this many hours.
            flag_artifact_hours: Flag when ``jump_duration + plummet_duration`` is at least this many hours.
            flag_proba_low: Lower bound of the uncertain predicted-probability interval (inclusive).
            flag_proba_high: Upper bound of the uncertain predicted-probability interval (inclusive).
        """
        print("\nAdding curve overlap detection to day features...")

        # Initialize curve overlap columns
        self.day_features['drinking_curve_overlap'] = 0
        self.day_features['total_curve_overlap_hours'] = 0.0

        # Always initialise 5 slots with full feature set
        _none_series = pd.Series([None] * len(self.day_features), index=self.day_features.index)
        for n in range(1, _MAX_CURVE_SLOTS + 1):
            self.day_features[f'curve_{n}_id'] = _none_series.copy()
            self.day_features[f'curve_{n}_overlap_hours'] = _none_series.copy()
            self.day_features[f'curve_{n}_high_quality_duration'] = _none_series.copy()
            self.day_features[f'curve_{n}_duration_CURVE'] = _none_series.copy()
            self.day_features[f'curve_{n}_extends_prior_day'] = _none_series.copy()
            self.day_features[f'curve_{n}_extends_next_day'] = _none_series.copy()
            for sfx in _PER_SLOT_CURVE_FEATURE_SUFFIXES:
                self.day_features[f'curve_{n}_{sfx}'] = _none_series.copy()
            self.day_features[f'curve_{n}_overlap_fraction'] = _none_series.copy()
            self.day_features[f'curve_{n}_start_offset_hours'] = _none_series.copy()
            self.day_features[f'curve_{n}_end_offset_hours'] = _none_series.copy()

        if curve_features_df is None or curve_features_df.empty:
            print("No curve features provided - curve overlap columns initialized with default values")
        else:
            required_curve_cols = [
                'subid',
                'dataset_id',
                'curve_id',
                'begin_CURVE',
                'end_CURVE',
                'high_quality_duration_CURVE',
            ]
            missing_curve_cols = [col for col in required_curve_cols if col not in curve_features_df.columns]
            if missing_curve_cols:
                print(f"Warning: Missing required columns in curve features: {missing_curve_cols}")
                return

            curve_features_df['begin_CURVE'] = pd.to_datetime(curve_features_df['begin_CURVE'])
            curve_features_df['end_CURVE'] = pd.to_datetime(curve_features_df['end_CURVE'])

            self.day_features['begin_day'] = pd.to_datetime(self.day_features['begin_day'])
            self.day_features['end_day'] = pd.to_datetime(self.day_features['end_day'])

            self.day_features = self.day_features.reset_index(drop=True)

            _c_sub = pd.to_numeric(curve_features_df['subid'], errors='coerce')
            _c_ds = pd.to_numeric(curve_features_df['dataset_id'], errors='coerce')

            # Resolve which extra suffixes are actually present in the curve DataFrame
            _available_suffixes = [s for s in _PER_SLOT_CURVE_FEATURE_SUFFIXES if s in curve_features_df.columns]

            for idx, day_row in self.day_features.iterrows():
                subid = day_row['SubID']
                dataset_id = day_row['Dataset_ID']
                day_start = day_row['begin_day']
                day_end = day_row['end_day']
                ds_num = pd.to_numeric(dataset_id, errors='coerce')
                sid_num = pd.to_numeric(subid, errors='coerce')

                subject_curves = curve_features_df[
                    (_c_sub == sid_num) & (_c_ds == ds_num)
                ]

                if subject_curves.empty:
                    continue

                overlapping_curves = subject_curves[
                    (subject_curves['begin_CURVE'] <= day_end) &
                    (subject_curves['end_CURVE'] >= day_start)
                ].copy()

                if overlapping_curves.empty:
                    continue

                # Sort by HQ duration descending; tie-break: higher AUC, then lower curve_id
                _hq = pd.to_numeric(overlapping_curves['high_quality_duration_CURVE'], errors='coerce').fillna(0)
                _auc = pd.to_numeric(overlapping_curves.get('auc_total_CURVE', 0), errors='coerce').fillna(0)
                _cid = pd.to_numeric(overlapping_curves['curve_id'], errors='coerce').fillna(0)
                overlapping_curves = overlapping_curves.assign(
                    _sort_hq=_hq, _sort_auc=_auc, _sort_cid=_cid
                ).sort_values(
                    ['_sort_hq', '_sort_auc', '_sort_cid'],
                    ascending=[False, False, True],
                ).drop(columns=['_sort_hq', '_sort_auc', '_sort_cid'])

                self.day_features.loc[day_row.name, 'drinking_curve_overlap'] = 1

                total_overlap = 0.0
                for curve_idx, curve_row in overlapping_curves.iterrows():
                    overlap_start = max(curve_row['begin_CURVE'], day_start)
                    overlap_end = min(curve_row['end_CURVE'], day_end)
                    overlap_hours = (overlap_end - overlap_start).total_seconds() / 3600
                    total_overlap += overlap_hours

                self.day_features.loc[day_row.name, 'total_curve_overlap_hours'] = total_overlap

                for n, (curve_idx, curve_row) in enumerate(overlapping_curves.iterrows(), 1):
                    if n > _MAX_CURVE_SLOTS:
                        break
                    overlap_start = max(curve_row['begin_CURVE'], day_start)
                    overlap_end = min(curve_row['end_CURVE'], day_end)
                    overlap_hours = (overlap_end - overlap_start).total_seconds() / 3600
                    extends_prior = curve_row['begin_CURVE'] < day_start
                    extends_next = curve_row['end_CURVE'] > day_end

                    self.day_features.loc[day_row.name, f'curve_{n}_id'] = curve_row['curve_id']
                    self.day_features.loc[day_row.name, f'curve_{n}_overlap_hours'] = overlap_hours
                    self.day_features.loc[day_row.name, f'curve_{n}_high_quality_duration'] = curve_row.get('high_quality_duration_CURVE', 0)
                    self.day_features.loc[day_row.name, f'curve_{n}_duration_CURVE'] = curve_row.get('duration_CURVE', np.nan)
                    self.day_features.loc[day_row.name, f'curve_{n}_extends_prior_day'] = int(extends_prior)
                    self.day_features.loc[day_row.name, f'curve_{n}_extends_next_day'] = int(extends_next)

                    for sfx in _available_suffixes:
                        self.day_features.loc[day_row.name, f'curve_{n}_{sfx}'] = curve_row.get(sfx, np.nan)

                    dur = pd.to_numeric(curve_row.get('duration_CURVE', np.nan), errors='coerce')
                    if pd.notna(dur) and dur > 0:
                        self.day_features.loc[day_row.name, f'curve_{n}_overlap_fraction'] = overlap_hours / dur
                    else:
                        self.day_features.loc[day_row.name, f'curve_{n}_overlap_fraction'] = np.nan

                    self.day_features.loc[day_row.name, f'curve_{n}_start_offset_hours'] = (
                        (curve_row['begin_CURVE'] - day_start).total_seconds() / 3600
                    )
                    self.day_features.loc[day_row.name, f'curve_{n}_end_offset_hours'] = (
                        (curve_row['end_CURVE'] - day_start).total_seconds() / 3600
                    )

            dayFeatures.assign_median_prior_from_all_qualifying_curves(self.day_features)
            dayFeatures.assign_median_prior_day_region_tac(self.day_features)

        from App.SDM.Machine_Learning.drinking_day_xgb import score_drinking_days
        _, missing_xgb = score_drinking_days(
            self.day_features,
            flag_non_wear_hours=flag_non_wear_hours,
            flag_very_negative_hours=flag_very_negative_hours,
            flag_artifact_hours=flag_artifact_hours,
            flag_proba_low=flag_proba_low,
            flag_proba_high=flag_proba_high,
        )
        if missing_xgb:
            print(
                f'Warning: {len(missing_xgb)} XGB drinking-day features were absent and scored as missing: '
                f'{missing_xgb[:12]}{"..." if len(missing_xgb) > 12 else ""}'
            )
        n_pos = int((self.day_features['pred_drinking_day'] == 1).sum())
        n_rel = int((self.day_features['prediction_reliable'] == 1).sum())
        print(
            f'XGB drinking-day scores: {n_pos} predicted drinking days, '
            f'{n_rel} / {len(self.day_features)} prediction_reliable.'
        )

        print(f"Curve overlap detection completed. Found curve overlaps on {self.day_features['drinking_curve_overlap'].sum()} days.")
        
        # Add curve overlap statistics to stats frames
        curve_stats = {
            'Days with curve overlap': self.day_features['drinking_curve_overlap'].sum(),
            'Total curve overlap hours across all days': self.day_features['total_curve_overlap_hours'].sum(),
            'Average curve overlap hours per overlapping day': self.day_features[self.day_features['drinking_curve_overlap'] == 1]['total_curve_overlap_hours'].mean() if self.day_features['drinking_curve_overlap'].sum() > 0 else 0,
            'Maximum curve overlap slots': _MAX_CURVE_SLOTS
        }
        
        curve_stats_df = pd.DataFrame.from_dict(curve_stats, orient='index', columns=['Value'])
        self.day_stat_frames.append(curve_stats_df)
        
        print("Curve overlap statistics added to stats frames.")

    def _skyn_metadata_table_for_morning(
        self,
        skyn_dates_csv_path=None,
        metadata_id_column='ID',
        metadata_burst_column='burst_id',
        first_day_column='first_day',
        last_date_column='last_day',
    ):
        """
        One row per (SubID, Dataset_ID) with burst date window and optional global study-day bounds.
        Prefers ``self.skyn_dates_metadata`` from ``filter_days_by_date_range``; otherwise reads CSV.
        """
        if self.skyn_dates_metadata is not None and len(self.skyn_dates_metadata):
            return self.skyn_dates_metadata.copy()
        if not skyn_dates_csv_path or not os.path.isfile(skyn_dates_csv_path):
            return None
        metadata = pd.read_csv(skyn_dates_csv_path)
        metadata.columns = metadata.columns.str.strip()
        required_meta_cols = [
            metadata_id_column,
            metadata_burst_column,
            first_day_column,
            last_date_column,
        ]
        if any(c not in metadata.columns for c in required_meta_cols):
            print(
                'Warning: skyn dates CSV missing required columns; '
                'cannot build skyn metadata for morning dates'
            )
            return None
        if metadata[metadata_id_column].isna().any():
            metadata[metadata_id_column] = metadata[metadata_id_column].ffill()
        metadata = metadata.dropna(subset=[metadata_id_column, metadata_burst_column])
        meta_attach_cols = list(required_meta_cols)
        for _opt in ('study_day_first_day', 'study_day_last_day'):
            if _opt in metadata.columns and _opt not in meta_attach_cols:
                meta_attach_cols.append(_opt)
        meta_attach = metadata[meta_attach_cols].drop_duplicates().copy()
        meta_rename = {
            metadata_id_column: 'SubID',
            metadata_burst_column: 'Dataset_ID',
            first_day_column: 'metadata_skyn_first_day',
            last_date_column: 'metadata_skyn_last_day',
        }
        if 'study_day_first_day' in meta_attach.columns:
            meta_rename['study_day_first_day'] = 'metadata_study_day_first'
        if 'study_day_last_day' in meta_attach.columns:
            meta_rename['study_day_last_day'] = 'metadata_study_day_last'
        meta_attach = meta_attach.rename(columns=meta_rename)
        meta_attach['SubID'] = pd.to_numeric(meta_attach['SubID'], errors='coerce')
        meta_attach['Dataset_ID'] = pd.to_numeric(meta_attach['Dataset_ID'], errors='coerce')
        meta_attach = meta_attach.dropna(subset=['SubID', 'Dataset_ID'])
        meta_attach['SubID'] = meta_attach['SubID'].astype(int)
        meta_attach['Dataset_ID'] = meta_attach['Dataset_ID'].astype(int)
        for _c in ('metadata_skyn_first_day', 'metadata_skyn_last_day'):
            meta_attach[_c] = pd.to_datetime(meta_attach[_c], errors='coerce')
        for _c in ('metadata_study_day_first', 'metadata_study_day_last'):
            if _c in meta_attach.columns:
                meta_attach[_c] = pd.to_numeric(meta_attach[_c], errors='coerce')
        return meta_attach

    @staticmethod
    def _expand_to_full_14_day_burst_grid(
        day_df: pd.DataFrame,
        skyn_meta: pd.DataFrame,
        n_days_per_burst: int = 14,
        day_start_hour: int = 6,
    ) -> pd.DataFrame:
        """
        Expand day-level rows to a complete 1..n_days_per_burst grid per (SubID, Dataset_ID),
        using the burst anchor date from skyn metadata (first_day / metadata_skyn_first_day).

        This is intended for exports where study days 1..14 should appear even when
        TAC data are missing for some days, so morning self-report can still merge onto those days.

        Requires skyn_meta to include: SubID, Dataset_ID, metadata_skyn_first_day.
        """
        if day_df is None or len(day_df) == 0:
            return day_df
        if skyn_meta is None or len(skyn_meta) == 0:
            return day_df
        needed = {'SubID', 'Dataset_ID', 'metadata_skyn_first_day'}
        if not needed.issubset(set(skyn_meta.columns)):
            return day_df

        out = day_df.copy()
        out['SubID'] = pd.to_numeric(out['SubID'], errors='coerce')
        out['Dataset_ID'] = pd.to_numeric(out['Dataset_ID'], errors='coerce')
        out['day_no'] = pd.to_numeric(out['day_no'], errors='coerce')
        out = out.dropna(subset=['SubID', 'Dataset_ID', 'day_no'])
        out['SubID'] = out['SubID'].astype(int)
        out['Dataset_ID'] = out['Dataset_ID'].astype(int)
        out['day_no'] = out['day_no'].astype(int)

        meta = skyn_meta[['SubID', 'Dataset_ID', 'metadata_skyn_first_day']].copy()
        meta['SubID'] = pd.to_numeric(meta['SubID'], errors='coerce')
        meta['Dataset_ID'] = pd.to_numeric(meta['Dataset_ID'], errors='coerce')
        meta = meta.dropna(subset=['SubID', 'Dataset_ID'])
        meta['SubID'] = meta['SubID'].astype(int)
        meta['Dataset_ID'] = meta['Dataset_ID'].astype(int)
        meta['metadata_skyn_first_day'] = pd.to_datetime(meta['metadata_skyn_first_day'], errors='coerce').dt.normalize()
        meta = meta.dropna(subset=['metadata_skyn_first_day']).drop_duplicates(subset=['SubID', 'Dataset_ID'])

        # Build scaffold rows: one per (SubID, Dataset_ID, day_no=1..N) with begin_day/end_day.
        scaff_parts = []
        for _, r in meta.iterrows():
            subid = int(r['SubID'])
            did = int(r['Dataset_ID'])
            d0 = r['metadata_skyn_first_day']
            day_nos = np.arange(1, int(n_days_per_burst) + 1, dtype=int)
            begin = d0 + pd.to_timedelta(day_nos - 1, unit='D') + pd.to_timedelta(int(day_start_hour), unit='h')
            end = begin + pd.to_timedelta(1, unit='D')
            scaff_parts.append(
                pd.DataFrame(
                    {
                        'SubID': subid,
                        'Dataset_ID': did,
                        'day_no': day_nos,
                        'begin_day': begin,
                        'end_day': end,
                    }
                )
            )
        if not scaff_parts:
            return out
        scaff = pd.concat(scaff_parts, ignore_index=True)

        # Outer join to retain full scaffold; preserve all observed day-level columns when present.
        merged = scaff.merge(out, on=['SubID', 'Dataset_ID', 'day_no'], how='left', suffixes=('_scaff', ''))

        # Fill begin/end from scaffold where missing; keep observed when present.
        for c in ('begin_day', 'end_day'):
            sc = f'{c}_scaff'
            if sc in merged.columns:
                if c in merged.columns:
                    merged[c] = merged[c].where(merged[c].notna(), merged[sc])
                else:
                    merged[c] = merged[sc]
                merged = merged.drop(columns=[sc], errors='ignore')

        # These scaffolded 1..14 days are *by definition* the in-window burst days from skyn metadata.
        # Fill inside_burst to 1 when missing so the export grids + Agreement_summary subset are consistent.
        if 'inside_burst' in merged.columns:
            merged['inside_burst'] = pd.to_numeric(merged['inside_burst'], errors='coerce')
            merged.loc[merged['inside_burst'].isna(), 'inside_burst'] = 1
        else:
            merged['inside_burst'] = 1

        # Mark rows that were synthetic (TAC-missing day rows) and set conservative defaults so they fall into unknown.
        merged['tac_day_missing'] = merged.get('tac_day_missing', np.nan)
        is_missing = merged['tac_day_missing'].isna() & merged.filter(regex='^begin_day$').notna() & merged.filter(regex='^SubID$').notna()
        # A row is "synthetic" if it had no observed values for a core TAC column like low_quality_percent and plot path.
        # Use presence of any non-null original columns besides scaffold keys to determine observed vs inserted.
        observed_mask = merged.drop(columns=['SubID', 'Dataset_ID', 'day_no', 'begin_day', 'end_day'], errors='ignore').notna().any(axis=1)
        synthetic = ~observed_mask
        merged.loc[synthetic, 'tac_day_missing'] = 1
        merged.loc[~synthetic, 'tac_day_missing'] = 0

        # Defaults for synthetic: non-drinking and unreliable, so agreement treats them as unknown.
        if 'pred_drinking_day' in merged.columns:
            merged.loc[synthetic, 'pred_drinking_day'] = 0
        if 'prediction_reliable' in merged.columns:
            merged.loc[synthetic, 'prediction_reliable'] = 0
        if 'low_quality_percent' in merged.columns:
            merged.loc[synthetic, 'low_quality_percent'] = 1.0
        if 'device_worn_percent_of_day' in merged.columns:
            merged.loc[synthetic, 'device_worn_percent_of_day'] = 0.0
        if 'device_worn_duration' in merged.columns:
            merged.loc[synthetic, 'device_worn_duration'] = 0.0
        if 'device_turned_on_duration' in merged.columns:
            merged.loc[synthetic, 'device_turned_on_duration'] = 0.0

        dayFeatures.assign_day_of_week_cat_from_begin_day(merged)
        return merged

    @classmethod
    def _coerce_morning_attach_value(cls, series, col):
        """Normalize skip tokens; coerce numeric-looking morning fields when possible."""
        s = series.copy()
        as_str = s.astype(str).str.strip()
        skip = as_str.str.upper().isin(
            {
                'NAN',
                'NONE',
                '',
                'SKIPPED',
                'CONDITION_SKIPPED',
                'NO_ANSWER',
                'NA',
                'N/A',
            }
        )
        s = s.where(~skip, np.nan)
        # Drink counts and 0/1 time bins are numeric; leave other types as-is.
        if col == 'mr_numdk' or col.startswith('mr_altim_'):
            return pd.to_numeric(s, errors='coerce')
        return s

    @classmethod
    def _prepare_morning_attach_columns(cls, m_df, extra_cols=None):
        """
        Ensure requested attach columns exist on a morning frame (NaN if absent in CSV).

        ``extra_cols`` defaults to empty: only ``morning_self_report_alcohol`` is merged unless
        a caller passes cohort-specific fields.
        """
        cols = list(extra_cols) if extra_cols else []
        if not cols:
            return []
        missing = []
        for c in cols:
            if c in m_df.columns:
                m_df[c] = cls._coerce_morning_attach_value(m_df[c], c)
            else:
                m_df[c] = np.nan
                missing.append(c)
        if missing:
            print(
                "  Morning attach columns missing from CSV (filled NaN on day rows): "
                + ", ".join(missing)
            )
            if {'mr_alst', 'mr_alfn'} & set(missing):
                print(
                    "  Note: protocol mr_alst/mr_alfn time pickers are not in morning.csv; "
                    "export uses mr_altim_1..7 time-bin checkboxes instead."
                )
        return cols

    @staticmethod
    def _merge_self_report_by_calendar_day(day_df, key_df):
        """Left-merge morning alcohol (+ optional QC fields) onto TAC days by calendar day."""
        value_cols = [
            c
            for c in key_df.columns
            if c not in ('SubID', 'Dataset_ID', 'morning_merge_date')
        ]
        k = key_df[['SubID', 'Dataset_ID', 'morning_merge_date'] + value_cols].copy()
        k['_date_key'] = pd.to_datetime(k['morning_merge_date'], errors='coerce').dt.strftime(
            '%Y-%m-%d'
        )
        k = k.drop(columns=['morning_merge_date'])
        out = day_df.copy()
        out['_date_key'] = pd.to_datetime(out['begin_day'], errors='coerce').dt.strftime('%Y-%m-%d')
        out = out.merge(k, on=['SubID', 'Dataset_ID', '_date_key'], how='left')
        out = out.drop(columns=['_date_key'], errors='ignore')
        out['morning_merge_date'] = pd.to_datetime(out['begin_day'], errors='coerce').dt.normalize()
        return out

    @staticmethod
    def _parse_morning_study_date_raw(series):
        dt = pd.to_datetime(series, errors='coerce')
        num = pd.to_numeric(series, errors='coerce')
        dt_excel = pd.to_datetime(num, unit='D', origin='1899-12-30', errors='coerce')
        use_excel = dt.isna() & num.notna() & (num >= 20000) & (num <= 60000)
        merged = dt.where(~use_excel, dt_excel)
        return pd.to_datetime(merged, errors='coerce').dt.normalize()

    @staticmethod
    def _looks_like_study_day(series):
        """Heuristic: mostly integers in 1..500 (study-day indices), not calendar strings."""
        num = pd.to_numeric(series, errors='coerce')
        if num.notna().sum() == 0:
            return False
        in_range = num.notna() & (num >= 1) & (num <= 500)
        return (in_range.sum() / max(1, num.notna().sum())) >= 0.8

    @staticmethod
    def _morning_self_report_series(m_df, self_report_col, response_type_col, submission_label):
        is_sub = (
            m_df[response_type_col].astype(str).str.strip() == submission_label
            if response_type_col in m_df.columns
            else pd.Series(True, index=m_df.index)
        )
        sr = pd.to_numeric(m_df[self_report_col], errors='coerce')
        out = pd.Series(np.nan, index=m_df.index, dtype=float)
        out.loc[is_sub & (sr == 1)] = 1.0
        out.loc[is_sub & (sr == 2)] = 0.0
        return out

    @staticmethod
    def _dedupe_morning_rows_by_priority(
        mcal,
        *,
        subset,
        response_type_col='Response Type',
        submission_label='Submission',
        self_report_col='mr_al_y',
        sr_coded_col='morning_self_report_alcohol',
    ):
        """
        One row per ``subset`` key with priority:
          1. Response Type == Submission over Missed / Partial / other
          2. Drinking yes (coded SR == 1 or raw self-report == 1) over no / missing
          3. Last row in file order among remaining ties
        """
        if mcal is None or len(mcal) == 0:
            return mcal
        out = mcal.copy()
        if response_type_col in out.columns:
            is_sub = out[response_type_col].astype(str).str.strip() == submission_label
        else:
            is_sub = pd.Series(True, index=out.index)
        out['_prio_submission'] = is_sub.astype(int)
        drink_yes = pd.Series(False, index=out.index)
        if sr_coded_col in out.columns:
            drink_yes |= pd.to_numeric(out[sr_coded_col], errors='coerce') == 1
        if self_report_col in out.columns:
            drink_yes |= pd.to_numeric(out[self_report_col], errors='coerce') == 1
        out['_prio_drink_yes'] = drink_yes.astype(int)
        if '_file_order' in out.columns:
            out['_prio_file'] = pd.to_numeric(out['_file_order'], errors='coerce')
        else:
            out['_prio_file'] = np.arange(len(out), dtype=float)
        out['_prio_file'] = out['_prio_file'].fillna(-1)
        sort_cols = list(subset) + ['_prio_submission', '_prio_drink_yes', '_prio_file']
        out = out.sort_values(sort_cols, kind='mergesort')
        out = out.drop_duplicates(subset=list(subset), keep='last')
        return out.drop(
            columns=['_prio_submission', '_prio_drink_yes', '_prio_file'],
            errors='ignore',
        )

    def _morning_export_add_window_qc(self, mcal_export, first_date_adjuster):
        """Add ``morning_implied_in_skyn_window`` from stored skyn metadata (same rule as TAC inside_burst)."""
        meta = self.skyn_dates_metadata
        if meta is None or not len(meta):
            return mcal_export
        need = [
            'SubID',
            'Dataset_ID',
            'metadata_skyn_first_day',
            'metadata_skyn_last_day',
        ]
        if not all(c in meta.columns for c in need):
            return mcal_export
        m = mcal_export.merge(meta[need], on=['SubID', 'Dataset_ID'], how='left')
        _qm = pd.to_datetime(m['morning_merge_date'], errors='coerce').dt.normalize()
        _w0 = pd.to_datetime(m['metadata_skyn_first_day'], errors='coerce') + pd.to_timedelta(
            int(first_date_adjuster), unit='D'
        )
        _w0 = pd.to_datetime(_w0, errors='coerce').dt.normalize()
        _w1 = pd.to_datetime(m['metadata_skyn_last_day'], errors='coerce').dt.normalize()
        _qh = m['metadata_skyn_first_day'].notna()
        _in_q = _qh & _qm.notna() & (_qm >= _w0) & (_qm <= _w1)
        _out_q = _qh & _qm.notna() & ~_in_q
        m['morning_implied_in_skyn_window'] = np.nan
        m.loc[_in_q, 'morning_implied_in_skyn_window'] = 1
        m.loc[_out_q, 'morning_implied_in_skyn_window'] = 0
        return m.drop(
            columns=['metadata_skyn_first_day', 'metadata_skyn_last_day'],
            errors='ignore',
        )

    @staticmethod
    def _compute_self_report_and_tac_comparison(
        df,
        pred_col='pred_drinking_day',
        sr_col='morning_self_report_alcohol',
    ):
        """
        Agreement labels when non-drinking days (``pred_col`` = 0) are trusted only when
        ``prediction_reliable`` = 1.

        ``prediction_reliable`` = 0 on a non-drinking day makes the agreement unknown
        (``true_unknown``, ``false_unknown``, or ``unknown_sr_missing``).

        ``true_unknown`` / ``false_unknown``: self-report yes/no but the day is not evaluable
        as non-drinking (inconclusive for TN/FN).

        ``unknown_sr_missing``: no drinking day, no merged morning self-report, and
        ``prediction_reliable`` is not 1, so agreement is **unknown** rather than
        ``negative_without_self_report``.
        """
        pred = pd.to_numeric(df[pred_col], errors='coerce').fillna(0) > 0
        if sr_col not in df.columns:
            sr = pd.Series(np.nan, index=df.index)
        else:
            sr = pd.to_numeric(df[sr_col], errors='coerce')
        sr_known = sr.notna()
        sr_yes = sr == 1
        sr_no = sr == 0

        if 'prediction_reliable' in df.columns:
            reliable = pd.to_numeric(df['prediction_reliable'], errors='coerce').fillna(0) == 1
        else:
            reliable = pd.Series(False, index=df.index)

        evaluable_non_drinking = reliable

        out = pd.Series('negative_without_self_report', index=df.index, dtype=object)
        out.loc[(~sr_known) & pred] = 'positive_without_self_report'
        out.loc[(~sr_known) & (~pred) & evaluable_non_drinking] = 'negative_without_self_report'
        out.loc[(~sr_known) & (~pred) & (~evaluable_non_drinking)] = 'unknown_sr_missing'
        out.loc[sr_known & pred & sr_yes] = 'true_positive'
        out.loc[sr_known & pred & sr_no] = 'false_positive'
        out.loc[sr_known & (~pred) & evaluable_non_drinking & sr_no] = 'true_negative'
        out.loc[sr_known & (~pred) & evaluable_non_drinking & sr_yes] = 'false_negative'
        out.loc[sr_known & (~pred) & (~evaluable_non_drinking) & sr_yes] = 'true_unknown'
        out.loc[sr_known & (~pred) & (~evaluable_non_drinking) & sr_no] = 'false_unknown'
        return out

    @staticmethod
    def _sensitivity_specificity_rows(tp, tn, fp, fn):
        """
        Self-report alcohol as reference: sensitivity = P(TAC+|SR+), specificity = P(TAC-|SR-),
        using only SR-known cells (TP, TN, FP, FN).
        """
        sens = (tp / (tp + fn)) if (tp + fn) > 0 else np.nan
        spec = (tn / (tn + fp)) if (tn + fp) > 0 else np.nan
        return pd.DataFrame(
            {
                'Category': [
                    'Sensitivity TP/(TP+FN) [reference: SR alcohol yes]',
                    'Specificity TN/(TN+FP) [reference: SR alcohol no]',
                ],
                'n': [sens, spec],
            }
        )

    @staticmethod
    def _agreement_label_grid_results(df_all, agreement_col):
        """
        3x3 grid (TAC status x self-report status) with each cell formatted as
        ``{burst label}: n (pct%)`` (same labels as Burst_* mini-tables), where pct is % of total day rows in df_all.
        """
        if df_all is None or len(df_all) == 0 or agreement_col not in df_all.columns:
            return pd.DataFrame()
        vh = df_all[agreement_col].astype(str).value_counts()
        denom_days = max(1, int(len(df_all)))
        burst_lbl = dict(dayFeatures.HQ_AGREEMENT_ORDER_AND_LABELS)

        def _cell(key: str) -> str:
            n = int(vh.get(key, 0))
            pct = 100.0 * n / denom_days
            label = burst_lbl.get(key, key)
            return f'{label}: {n} ({pct:.1f}%)'

        return pd.DataFrame(
            [
                {
                    'TAC \\ self-report': 'Positive',
                    'SR true (alcohol yes)': _cell('true_positive'),
                    'SR false (alcohol no)': _cell('false_positive'),
                    'SR missing': _cell('positive_without_self_report'),
                },
                {
                    'TAC \\ self-report': 'Negative (sufficient data)',
                    'SR true (alcohol yes)': _cell('false_negative'),
                    'SR false (alcohol no)': _cell('true_negative'),
                    'SR missing': _cell('negative_without_self_report'),
                },
                {
                    'TAC \\ self-report': 'Unknown',
                    'SR true (alcohol yes)': _cell('true_unknown'),
                    'SR false (alcohol no)': _cell('false_unknown'),
                    'SR missing': _cell('unknown_sr_missing'),
                },
            ]
        )

    @staticmethod
    def _burst_timeframe_inclusion_rows(df_full):
        """
        Count Skyn day rows included vs excluded by burst/wear window (``inside_burst``).
        Returns empty DataFrame if ``inside_burst`` is absent.
        """
        if df_full is None or len(df_full) == 0 or 'inside_burst' not in df_full.columns:
            return pd.DataFrame(columns=['Category', 'n'])
        ib = pd.to_numeric(df_full['inside_burst'], errors='coerce')
        n_inc = int((ib == 1).sum())
        n_exc = int((ib == 0).sum())
        n_mis = int(ib.isna().sum())
        rows = [
            ('— Skyn day rows vs burst / wear window —', np.nan),
            ('Included (inside_burst = 1)', n_inc),
            ('Excluded (inside_burst = 0)', n_exc),
        ]
        if n_mis:
            rows.append(('inside_burst not coded (missing)', n_mis))
        return pd.DataFrame(rows, columns=['Category', 'n'])

    @staticmethod
    def _morning_tac_summary_blocks(
        df_all,
        pred_col,
        hq_agreement_col,
        df_full_burst_counts=None,
    ):
        """Build DataFrames for Agreement_summary / Stats: burst counts, day totals, and agreement labels."""
        if pred_col not in df_all.columns:
            return pd.DataFrame(
                {
                    'Category': [f"'{pred_col}' not in subset — cannot summarize morning vs TAC"],
                    'n': [np.nan],
                }
            )
        pred = pd.to_numeric(df_all[pred_col], errors='coerce').fillna(0) > 0
        sr = (
            pd.to_numeric(df_all['morning_self_report_alcohol'], errors='coerce')
            if 'morning_self_report_alcohol' in df_all.columns
            else pd.Series(np.nan, index=df_all.index)
        )

        # TAC-based day type counts: a non-drinking day is evaluable only when prediction_reliable = 1.
        if 'prediction_reliable' in df_all.columns:
            reliable = pd.to_numeric(df_all['prediction_reliable'], errors='coerce').fillna(0) == 1
        else:
            reliable = pd.Series(False, index=df_all.index)
        evaluable_non_drinking = reliable
        tac_unknown_non_drinking = (~pred) & (~evaluable_non_drinking)

        totals = pd.DataFrame(
            {
                'Category': [
                    'Self-report: alcohol yes',
                    'Self-report: alcohol no',
                    'Self-report: missing',
                    'TAC drinking day: positive',
                    'TAC non-drinking day: negative (sufficient data)',
                    'TAC non-drinking day: unknown (insufficient data)',
                ],
                'n': [
                    int((sr == 1).sum()),
                    int((sr == 0).sum()),
                    int(sr.isna().sum()),
                    int(pred.sum()),
                    int(((~pred) & evaluable_non_drinking).sum()),
                    int(tac_unknown_non_drinking.sum()),
                ],
            }
        )

        # Participants with ≥1 drinking day by burst (MAPTAC-style summary; mirrors evaluate_day_output participant stats).
        drinking_day_participant_rows = []
        if 'SubID' in df_all.columns and 'Dataset_ID' in df_all.columns and pred_col in df_all.columns:
            work = df_all[['SubID', 'Dataset_ID', pred_col]].copy()
            work['SubID'] = pd.to_numeric(work['SubID'], errors='coerce')
            work['Dataset_ID'] = pd.to_numeric(work['Dataset_ID'], errors='coerce')
            work = work.dropna(subset=['SubID', 'Dataset_ID'])
            work['_is_drinking_day'] = (pd.to_numeric(work[pred_col], errors='coerce').fillna(0) > 0).astype(int)

            per_participant_burst = (
                work.groupby(['SubID', 'Dataset_ID'], dropna=False)['_is_drinking_day']
                .sum()
                .reset_index()
                .rename(columns={'_is_drinking_day': 'drinking_days'})
            )
            # Denominator for each burst: participants with any day row in that burst.
            for b in (1, 2, 3):
                in_b = per_participant_burst['Dataset_ID'] == b
                denom = int(per_participant_burst.loc[in_b, 'SubID'].nunique())
                n_with_drinking = int(per_participant_burst.loc[in_b & (per_participant_burst['drinking_days'] > 0), 'SubID'].nunique())
                pct = 0.0 if denom == 0 else (100.0 * n_with_drinking / denom)
                drinking_day_participant_rows.append(
                    {
                        'Category': f'Burst {b}: participants with ≥1 drinking day',
                        'n': f'{n_with_drinking} ({pct:.1f}%) of {denom}',
                    }
                )
            denom_total = int(per_participant_burst['SubID'].nunique())
            n_any = int(per_participant_burst.loc[per_participant_burst['drinking_days'] > 0, 'SubID'].nunique())
            pct_any = 0.0 if denom_total == 0 else (100.0 * n_any / denom_total)
            drinking_day_participant_rows.append(
                {
                    'Category': 'Overall (bursts 1–3): participants with ≥1 drinking day',
                    'n': f'{n_any} ({pct_any:.1f}%) of {denom_total}',
                }
            )
        drinking_day_participant_df = (
            pd.DataFrame(drinking_day_participant_rows, columns=['Category', 'n'])
            if drinking_day_participant_rows
            else pd.DataFrame(columns=['Category', 'n'])
        )

        # Average 14-day burst composition: drinking vs reliable non-drinking vs unknown (unreliable or missing days).
        burst_breakdown_rows = []
        if (
            'SubID' in df_all.columns
            and 'Dataset_ID' in df_all.columns
            and pred_col in df_all.columns
        ):
            bd_cols = ['SubID', 'Dataset_ID', pred_col]
            if 'prediction_reliable' in df_all.columns:
                bd_cols.append('prediction_reliable')
            bd = df_all[bd_cols].copy()
            bd['SubID'] = pd.to_numeric(bd['SubID'], errors='coerce')
            bd['Dataset_ID'] = pd.to_numeric(bd['Dataset_ID'], errors='coerce')
            bd = bd.dropna(subset=['SubID', 'Dataset_ID'])

            drinking = pd.to_numeric(bd[pred_col], errors='coerce').fillna(0) > 0
            if 'prediction_reliable' in bd.columns:
                reliable_bd = pd.to_numeric(bd['prediction_reliable'], errors='coerce').fillna(0) == 1
            else:
                reliable_bd = pd.Series(False, index=bd.index)

            bd['_drinking_day'] = drinking.astype(int)
            bd['_valid_non_drinking'] = ((~drinking) & reliable_bd).astype(int)
            bd['_below_quality_data'] = ((~drinking) & (~reliable_bd)).astype(int)

            agg = (
                bd.groupby(['SubID', 'Dataset_ID'])
                .agg(
                    _drinking_day=('_drinking_day', 'sum'),
                    _valid_non_drinking=('_valid_non_drinking', 'sum'),
                    _below_quality_data=('_below_quality_data', 'sum'),
                    _n_days=('SubID', 'count'),
                )
                .reset_index()
            )
            n_days_per_burst = 14.0
            agg['_missing'] = (n_days_per_burst - agg['_n_days']).clip(lower=0)
            agg['_unknown'] = agg['_below_quality_data'] + agg['_missing']

            if len(agg):
                mean_drink = float(agg['_drinking_day'].mean())
                mean_valid = float(agg['_valid_non_drinking'].mean())
                mean_unknown = float(agg['_unknown'].mean())
                pct_drink = (mean_drink / n_days_per_burst) * 100.0
                pct_valid = (mean_valid / n_days_per_burst) * 100.0
                pct_unknown = (mean_unknown / n_days_per_burst) * 100.0

                burst_breakdown_rows.extend(
                    [
                        {
                            'Category': 'Avg drinking days per 14-day burst (pred_col > 0; quality not applied)',
                            'n': f'{mean_drink:.1f} ({pct_drink:.0f}%)',
                        },
                        {
                            'Category': 'Avg non-drinking days per 14-day burst (prediction_reliable = 1)',
                            'n': f'{mean_valid:.1f} ({pct_valid:.0f}%)',
                        },
                        {
                            'Category': 'Avg unknown days per 14-day burst (non-drinking days with prediction_reliable = 0 OR missing day)',
                            'n': f'{mean_unknown:.1f} ({pct_unknown:.0f}%)',
                        },
                    ]
                )
        burst_breakdown_df = (
            pd.DataFrame(burst_breakdown_rows, columns=['Category', 'n'])
            if burst_breakdown_rows
            else pd.DataFrame(columns=['Category', 'n'])
        )

        # Sensitivity/specificity is computed from SR-known TP/TN/FP/FN.
        if hq_agreement_col in df_all.columns:
            vh = df_all[hq_agreement_col].astype(str).value_counts()
            tp_h = int(vh.get('true_positive', 0))
            tn_h = int(vh.get('true_negative', 0))
            fp_h = int(vh.get('false_positive', 0))
            fn_h = int(vh.get('false_negative', 0))
            sens_spec_hq = dayFeatures._sensitivity_specificity_rows(tp_h, tn_h, fp_h, fn_h)
        else:
            sens_spec_hq = pd.DataFrame(columns=['Category', 'n'])

        spacer = pd.DataFrame({'Category': [''], 'n': [np.nan]})
        sec_tot = pd.DataFrame({'Category': ['— Day-level totals (SR / TAC) —'], 'n': [np.nan]})
        head_parts = []
        head_parts.extend(
            [
                sec_tot,
                pd.DataFrame({'Category': ['Total Study Days'], 'n': [len(df_all)]}),
                totals,
                spacer,
                pd.DataFrame({'Category': ['— Participants with ≥1 drinking day (Based on TAC only) —'], 'n': [np.nan]}),
                drinking_day_participant_df,
                spacer,
                pd.DataFrame({'Category': ['— Average 14-day burst composition (Based on TAC only) —'], 'n': [np.nan]}),
                burst_breakdown_df,
                spacer,
                pd.DataFrame({'Category': ['— Unknown breakdown (Based on TAC only) —'], 'n': [np.nan]}),
                (
                    pd.DataFrame(
                        [
                            {
                                'Category': 'Unknown due to prediction_reliable = 0',
                                'n': int((df_all.get('unknown_reason', '') == 'prediction_unreliable').sum()),
                            },
                            {
                                'Category': 'Unknown due to missing Skyn day',
                                'n': int((df_all.get('unknown_reason', '') == 'missing_day_expanded').sum()),
                            },
                        ]
                    )
                    if 'unknown_reason' in df_all.columns
                    else pd.DataFrame({'Category': ['(unknown_reason not available)'], 'n': [np.nan]})
                ),
                spacer,
                pd.DataFrame(
                    {
                        'Category': [
                            'Note: Unknown days are non-drinking days with prediction_reliable = 0, '
                            'or days with no Skyn TAC row.',
                        ],
                        'n': [np.nan],
                    }
                ),
            ]
        )
        return pd.concat(head_parts, ignore_index=True)

    @staticmethod
    def build_tac_self_report_methodology_dataframe(pred_col='pred_drinking_day'):
        """
        User-facing explanations for the TAC + morning self-report validation workbook:
        column names, definitions, and how values are derived (no code references).
        """
        pc = pred_col
        if pc == 'pred_drinking_day':
            detection_text = (
                f'Column: ``{pc}`` — 1 = drinking day, 0 = not.\n\n'
                'This column is the 08.05 XGBoost drinking-day model (probability at least 0.5). '
                'It is scored for every day after curve slots are attached, including days with no TAC. '
                '``proba_drinking_day`` is the model probability. ``prediction_reliable`` is a separate flag: '
                'a non-drinking day is a true or false negative only when that flag is 1.'
            )
        else:
            detection_text = (
                f'Column: ``{pc}`` — 1 = drinking day, 0 = not.\n\n'
                'Primary rule: The day is 1 if at least one predicted-drinking curve '
                '(``DRINKING_PRED`` = 1 in curve features) has its **start** in that calendar day: '
                '``begin_CURVE`` ≥ ``begin_day`` and ``begin_CURVE`` < ``end_day``. '
                'Minimum duration, ``REGION_VALID``, below-threshold, rise-phase, and other curve QC '
                'are applied **when curves are built** (they determine whether ``DRINKING_PRED`` is 1); '
                'the day-level overlap step only checks overlap with the day window and whether the curve start '
                'falls inside it.\n\n'
                'Secondary rule (OR): When ``above_threshold_percent_of_day`` and '
                '``above_threshold_high_quality_percent`` are present on the day row, the day is also set to 1 '
                'if ``above_threshold_percent_of_day`` ≥ 0.75 **and** '
                '``above_threshold_high_quality_percent`` > 0.60. '
                'The first is the fraction of the day with TAC at or above the curve threshold; the second is '
                'the fraction of **above-threshold** time that is high quality. '
                'This can flag a drinking day even when no predicted-drinking curve starts on that day.'
            )
        rows = [
            (
                'Drinking day detection',
                detection_text,
            ),
            (
                'Non-drinking vs sufficient-data non-drinking days',
                (
                    f'Non-drinking day: ``{pc}`` = 0.\n\n'
                    'Sufficient-data non-drinking agreement: ``self_report_and_tac_comparison`` assigns true/false '
                    f'negative only when ``{pc}`` = 0 and ``prediction_reliable`` = 1.\n\n'
                    'Unreliable prediction: if ``prediction_reliable`` = 0, a non-drinking day is '
                    'unknown (not a true or false negative).\n\n'
                    f'If ``{pc}`` = 0 and ``prediction_reliable`` = 0, labels ``true_unknown`` or '
                    '``false_unknown`` are used when self-report is present.\n\n'
                    'Agreement column on ``Morning_day_merge`` is ``self_report_and_tac_comparison`` '
                    '(same labels as Agreement_summary).'
                ),
            ),
            (
                'Unreliable predictions and unknown labels (non-drinking days)',
                (
                    '``prediction_reliable`` is 1 only when none of these fire: 6 or more hours not worn '
                    '(24 minus wear time), 1 or more hours of TAC below −15, 1 or more hours of jumps plus plummets, '
                    'or a drinking probability inside 0.05–0.95.\n\n'
                    'Agreement use: drinking days (``{pc}`` = 1) are never set aside for reliability. For '
                    'non-drinking days (``{pc}`` = 0), TN/FN only when ``prediction_reliable`` = 1. Otherwise '
                    '(with self-report present) labels are ``true_unknown`` (SR yes) or ``false_unknown`` (SR no).\n\n'
                    '``unknown_sr_missing``: non-drinking day (``{pc}`` = 0), morning self-report missing, and '
                    '``prediction_reliable`` is not 1. The row is unknown.\n\n'
                    'A missing ``prediction_reliable`` value is treated as not reliable.'
                ).format(pc=pc),
            ),
            (
                'Morning self-report on each day row',
                (
                    'Columns: ``morning_self_report_alcohol`` — binary alcohol yes/no from the morning report for '
                    'that day when a match exists; ``morning_report_matched`` — indicates a non-missing merged '
                    'value.\n\n'
                    'How rows are matched: when burst date windows are available, each morning report is aligned '
                    'to the TAC day using SubID, Dataset_ID, and the calendar date of the day '
                    '(``begin_day``), consistent with the study wear window. Otherwise the match uses **SubID**, '
                    '**Dataset_ID**, and **day_no** (study day index within the burst).\n\n'
                    'Optional: ``morning_implied_in_skyn_window`` may appear on annotated morning tables to show '
                    'whether the implied calendar date falls inside the skyn wear window.'
                ),
            ),
            (
                'Agreement columns and workbook subsets',
                (
                    '``self_report_and_tac_comparison`` (Morning_day_merge, Agreement_summary, burst mini-tables, plot labels): '
                    'cross-classifies ``morning_self_report_alcohol`` vs ``{pc}``. '
                    'For non-drinking days, TN/FN only when ``prediction_reliable`` = 1; '
                    'otherwise ``true_unknown`` / ``false_unknown`` when self-report is present. If morning '
                    'self-report is missing on a non-drinking day: ``unknown_sr_missing`` (unknown [SR missing] on '
                    'plots) when ``prediction_reliable`` is not 1; otherwise Negative TAC, no morning SR when the day '
                    'is a reliable non-drinking day.\n\n'
                    'Where ``inside_burst`` is used: Agreement_summary and the Burst_* sheets use only rows '
                    'with ``inside_burst`` = 1 when that column exists; otherwise all day rows are used.\n\n'
                    'Burst_* sheets: participant columns are ordered left to right from fewer to more drinking '
                    'days (days with ``{pc}`` = 1 in that burst; ties keep the original groupby order).'
                ).format(pc=pc),
            ),
        ]
        return pd.DataFrame(rows, columns=['Topic', 'Explanation'])

    @staticmethod
    def build_tac_self_report_variable_key_dataframe(pred_col='pred_drinking_day'):
        """
        Stacked tables for the TAC + morning workbook ``Variable_key`` sheet: morning agreement
        columns, then curated day-level definitions from ``report_guide.day_feature_descriptions``.
        """
        from App.SDM.Documenting.report_guide import report_guide

        dfd = report_guide.day_feature_descriptions
        morning_keys = [
            'morning_report_matched',
            'self_report_and_tac_comparison',
        ]
        rows_m = [{'Variable': k, 'Description': dfd[k]} for k in morning_keys if k in dfd]

        day_keys = [
            'SubID',
            'Dataset_ID',
            'day_no',
            'begin_day',
            'day_of_week_cat',
            'end_day',
            'inside_burst',
            'tac_day_missing',
            'low_quality_percent',
            'device_worn_percent_of_day',
            pred_col,
            'drinking_curve_overlap',
            'above_threshold_percent_of_day',
            'above_threshold_high_quality_percent',
            'below_threshold_percent',
            'total_curve_overlap_hours',
        ]
        rows_d = [{'Variable': k, 'Description': dfd[k]} for k in day_keys if k in dfd]

        hdr_m = pd.DataFrame(
            [{'Variable': 'Morning merge and agreement columns', 'Description': ''}],
        )
        hdr_d = pd.DataFrame(
            [
                {
                    'Variable': 'Day-level columns (report guide)',
                    'Description': 'Selected fields from the project day feature definitions.',
                }
            ],
        )
        gap = pd.DataFrame([{'Variable': '', 'Description': ''}])
        return pd.concat(
            [
                hdr_m,
                pd.DataFrame(rows_m),
                gap,
                hdr_d,
                pd.DataFrame(rows_d),
            ],
            ignore_index=True,
        )

    @staticmethod
    def agreement_burst_cell_display_label(agreement_value):
        """
        Compact agreement text above each burst-grid plot from ``self_report_and_tac_comparison``.
        """
        raw = agreement_value
        if raw is None or pd.isna(raw) or (isinstance(raw, str) and not str(raw).strip()):
            return ''
        s = str(raw).strip()
        m = {
            'true_positive': 'TP',
            'true_negative': 'TN',
            'false_positive': 'FP',
            'false_negative': 'FN',
            'true_unknown': 'True unknown',
            'false_unknown': 'False unknown',
            'unknown_sr_missing': 'unknown [SR missing]',
            'positive_without_self_report': 'Positive TAC, no morning SR',
            'negative_without_self_report': 'Negative TAC, no morning SR',
        }
        return m.get(s, s)

    @staticmethod
    def filter_day_features_by_subid_range(df, subid_min=None, subid_max=None):
        """Restrict day-level rows to ``subid_min`` <= SubID <= ``subid_max`` (inclusive)."""
        if subid_min is None and subid_max is None:
            return df
        if 'SubID' not in df.columns:
            print('  Warning: SubID column missing; subid range filter skipped')
            return df
        sub = pd.to_numeric(df['SubID'], errors='coerce')
        mask = sub.notna()
        if subid_min is not None:
            mask &= sub >= int(subid_min)
        if subid_max is not None:
            mask &= sub <= int(subid_max)
        out = df.loc[mask].copy()
        lo = subid_min if subid_min is not None else '…'
        hi = subid_max if subid_max is not None else '…'
        print(f'  SubID filter [{lo}, {hi}]: {len(out)} / {len(df)} rows')
        return out

    @staticmethod
    def self_report_plot_cell_label(day_row):
        """
        Morning self-report text for the plot cell (under the embedded image).

        Returns ``self-report: drinking``, ``self-report: no drinking``, or
        ``self-report: missing`` from ``morning_self_report_alcohol`` only.
        """
        if day_row is None or (hasattr(day_row, 'empty') and day_row.empty):
            return 'self-report: missing'
        if isinstance(day_row, pd.DataFrame):
            if day_row.empty:
                return 'self-report: missing'
            row = day_row.iloc[0]
        else:
            row = day_row
        if 'morning_self_report_alcohol' not in row.index:
            return 'self-report: missing'
        sr = pd.to_numeric(row['morning_self_report_alcohol'], errors='coerce')
        if pd.isna(sr):
            return 'self-report: missing'
        if int(sr) == 1:
            return 'self-report: drinking'
        if int(sr) == 0:
            return 'self-report: no drinking'
        return 'self-report: missing'

    def _append_morning_tac_summary_stats(self, pred_col='pred_drinking_day'):
        """Append morning/TAC summary tables to ``day_stat_frames`` for main day Stats sheet."""
        if 'self_report_and_tac_comparison' not in self.day_features.columns:
            return
        df = self.day_features
        if 'inside_burst' in df.columns:
            sub = df[df['inside_burst'] == 1].copy()
            hdr = 'Morning vs TAC summary (inside_burst == 1)'
        else:
            sub = df.copy()
            hdr = 'Morning vs TAC summary (all day rows)'
        blk = self._morning_tac_summary_blocks(
            sub,
            pred_col=pred_col,
            hq_agreement_col='self_report_and_tac_comparison',
            df_full_burst_counts=df,
        )
        out = pd.concat(
            [pd.DataFrame({'Label': [hdr], 'Value': [np.nan]}), blk.rename(columns={'Category': 'Label', 'n': 'Value'})],
            ignore_index=True,
        )
        self.day_stat_frames.append(out)

    def prepare_morning_self_report_for_tac_merge(
        self,
        morning_csv_path,
        self_report_col='mr_al_y',
        trigger_col='Trigger Name',
        id_col='id',
        response_type_col='Response Type',
        submission_label='Submission',
        morning_study_date_column=None,
        skyn_dates_csv_path=None,
        metadata_id_column='ID',
        metadata_burst_column='burst_id',
        first_day_column='first_day',
        last_date_column='last_day',
        first_date_adjuster=-1,
        extra_attach_cols=None,
    ):
        """
        Load ``morning.csv``, resolve each row's calendar **study_date** (``morning_merge_date``) using
        ``self.skyn_dates_metadata`` when present (else ``skyn_dates_csv_path``), and build merge keys.

        Sets:
            ``self.morning_merge_key_calendar``: ``SubID``, ``Dataset_ID``, ``morning_merge_date``,
            ``morning_self_report_alcohol`` (preferred join to TAC ``begin_day``), plus any
            cohort-specific QC fields listed in ``extra_attach_cols`` (default: none).
            ``self.morning_merge_key_day_no``: legacy ``day_no`` join when calendar resolution is unavailable.
            ``self.morning_annotated``: long-form morning rows with dates + optional QC column for export.
        """
        self.morning_merge_key_calendar = None
        self.morning_merge_key_day_no = None
        self.morning_annotated = None
        # Empty by default: cohort scripts pass QC fields via extra_attach_cols.
        self._morning_extra_attach_cols = list(extra_attach_cols or ())

        if not os.path.isfile(morning_csv_path):
            print(f"Warning: Morning report file not found: {morning_csv_path}")
            return

        morning = pd.read_csv(morning_csv_path)
        morning.columns = morning.columns.str.strip()

        if id_col not in morning.columns or self_report_col not in morning.columns:
            print(f"Warning: morning CSV missing required columns ({id_col} / {self_report_col})")
            return

        if trigger_col not in morning.columns:
            print(f"Warning: morning CSV missing '{trigger_col}' for burst parsing")
            return

        def _burst_from_trigger(val):
            if pd.isna(val):
                return np.nan
            m = re.search(r'Burst\s+(\d+)', str(val), flags=re.IGNORECASE)
            return int(m.group(1)) if m else np.nan

        morning['_file_order'] = morning.index
        morning['Dataset_ID'] = morning[trigger_col].map(_burst_from_trigger)
        morning['SubID'] = pd.to_numeric(morning[id_col], errors='coerce')
        morning = morning.dropna(subset=['SubID', 'Dataset_ID']).copy()
        morning['SubID'] = morning['SubID'].astype(int)
        morning['Dataset_ID'] = morning['Dataset_ID'].astype(int)
        morning_sorted = morning.sort_values(['SubID', 'Dataset_ID', '_file_order'])

        mm = self._skyn_metadata_table_for_morning(
            skyn_dates_csv_path=skyn_dates_csv_path,
            metadata_id_column=metadata_id_column,
            metadata_burst_column=metadata_burst_column,
            first_day_column=first_day_column,
            last_date_column=last_date_column,
        )

        def _finalize_calendar_key(mcal, label):
            mcal = mcal.dropna(subset=['morning_merge_date']).copy()
            if mcal.empty:
                print(f"  Warning: no morning rows left after date resolution ({label})")
                return False
            mcal['morning_self_report_alcohol'] = self._morning_self_report_series(
                mcal, self_report_col, response_type_col, submission_label
            )
            attach_cols = self._prepare_morning_attach_columns(
                mcal, self._morning_extra_attach_cols
            )
            mcal = self._dedupe_morning_rows_by_priority(
                mcal,
                subset=['SubID', 'Dataset_ID', 'morning_merge_date'],
                response_type_col=response_type_col,
                submission_label=submission_label,
                self_report_col=self_report_col,
                sr_coded_col='morning_self_report_alcohol',
            )
            keep = [
                'SubID',
                'Dataset_ID',
                'morning_merge_date',
                'morning_self_report_alcohol',
            ] + attach_cols
            self.morning_merge_key_calendar = mcal[keep].copy()
            for c, dtype in (
                ('SubID', np.int64),
                ('Dataset_ID', np.int64),
            ):
                self.morning_merge_key_calendar[c] = pd.to_numeric(
                    self.morning_merge_key_calendar[c], errors='coerce'
                ).astype(dtype)
            self.morning_merge_key_calendar['morning_merge_date'] = pd.to_datetime(
                self.morning_merge_key_calendar['morning_merge_date'], errors='coerce'
            ).dt.normalize()
            print(f"  {label}")
            if attach_cols:
                print(f"  Attach QC morning columns on calendar key: {', '.join(attach_cols)}")
            return True

        # --- Path 1: explicit date / study-day column ---
        if morning_study_date_column and morning_study_date_column in morning_sorted.columns:
            if 'begin_day' not in self.day_features.columns:
                print(
                    "Warning: day_features missing 'begin_day' for calendar morning merge; "
                    "will try slot or day_no if needed"
                )
            mcal = morning_sorted.copy()
            _sd_raw = mcal[morning_study_date_column]
            # Integer study days must not go through pd.to_datetime: small ints become ~1970-01-01 and
            # skip the skyn metadata mapping entirely.
            if self._looks_like_study_day(_sd_raw):
                if mm is None:
                    print(
                        f"  Warning: '{morning_study_date_column}' looks like study_day but "
                        "no skyn metadata (run filter_days_by_date_range or pass skyn_dates_csv_path)"
                    )
                    mcal['morning_merge_date'] = pd.NaT
                else:
                    mcal = mcal.merge(mm, on=['SubID', 'Dataset_ID'], how='left')
                    sd = pd.to_numeric(mcal[morning_study_date_column], errors='coerce')
                    has_global = (
                        'metadata_study_day_first' in mcal.columns
                        and mcal['metadata_study_day_first'].notna().any()
                    )
                    if has_global:
                        in_global = (
                            mcal['metadata_study_day_first'].notna()
                            & mcal['metadata_study_day_last'].notna()
                            & (sd >= mcal['metadata_study_day_first'])
                            & (sd <= mcal['metadata_study_day_last'])
                        )
                        day_offset = (sd - mcal['metadata_study_day_first']).where(
                            in_global, sd - 1
                        )
                    else:
                        day_offset = sd - 1
                    mcal['morning_merge_date'] = (
                        mcal['metadata_skyn_first_day']
                        + pd.to_timedelta(int(first_date_adjuster), unit='D')
                        + pd.to_timedelta(day_offset, unit='D')
                    )
                    mcal['morning_merge_date'] = pd.to_datetime(
                        mcal['morning_merge_date'], errors='coerce'
                    ).dt.normalize()
                    mcal = mcal.drop(
                        columns=[
                            c
                            for c in (
                                'metadata_skyn_first_day',
                                'metadata_skyn_last_day',
                                'metadata_study_day_first',
                                'metadata_study_day_last',
                            )
                            if c in mcal.columns
                        ],
                        errors='ignore',
                    )
                    print(
                        f"  Integer '{morning_study_date_column}' → calendar date via skyn metadata "
                        f"({first_day_column} + {first_date_adjuster}; global bounds when present)."
                    )
            else:
                mcal['morning_merge_date'] = self._parse_morning_study_date_raw(_sd_raw)
                if mcal['morning_merge_date'].isna().all() and self._looks_like_study_day(_sd_raw):
                    if mm is None:
                        print(
                            f"  Warning: '{morning_study_date_column}' could not parse as dates and "
                            "needs skyn metadata for study_day"
                        )
                    else:
                        mcal = mcal.merge(mm, on=['SubID', 'Dataset_ID'], how='left')
                        sd = pd.to_numeric(mcal[morning_study_date_column], errors='coerce')
                        has_global = (
                            'metadata_study_day_first' in mcal.columns
                            and mcal['metadata_study_day_first'].notna().any()
                        )
                        if has_global:
                            in_global = (
                                mcal['metadata_study_day_first'].notna()
                                & mcal['metadata_study_day_last'].notna()
                                & (sd >= mcal['metadata_study_day_first'])
                                & (sd <= mcal['metadata_study_day_last'])
                            )
                            day_offset = (sd - mcal['metadata_study_day_first']).where(
                                in_global, sd - 1
                            )
                        else:
                            day_offset = sd - 1
                        mcal['morning_merge_date'] = (
                            mcal['metadata_skyn_first_day']
                            + pd.to_timedelta(int(first_date_adjuster), unit='D')
                            + pd.to_timedelta(day_offset, unit='D')
                        )
                        mcal['morning_merge_date'] = pd.to_datetime(
                            mcal['morning_merge_date'], errors='coerce'
                        ).dt.normalize()
                        mcal = mcal.drop(
                            columns=[
                                c
                                for c in (
                                    'metadata_skyn_first_day',
                                    'metadata_skyn_last_day',
                                    'metadata_study_day_first',
                                    'metadata_study_day_last',
                                )
                                if c in mcal.columns
                            ],
                            errors='ignore',
                        )
                        print(
                            f"  Unparseable values in '{morning_study_date_column}' → "
                            f"mapped via skyn study_day columns (+{first_date_adjuster} d anchor)."
                        )

            mcal_export = mcal.copy()
            mcal_export = self._morning_export_add_window_qc(mcal_export, first_date_adjuster)
            _mdt = pd.to_datetime(mcal_export['morning_merge_date'], errors='coerce')
            mcal_export['morning_calendar_date'] = _mdt.dt.strftime('%Y-%m-%d')
            mcal_export.loc[_mdt.isna(), 'morning_calendar_date'] = np.nan
            self.morning_annotated = mcal_export.drop(
                columns=[c for c in mcal_export.columns if c.startswith('_')],
                errors='ignore',
            )

            n_bad = int(mcal['morning_merge_date'].isna().sum())
            if n_bad:
                print(
                    f"  Warning: {n_bad} morning rows invalid/missing "
                    f"'{morning_study_date_column}' after parsing; excluded from merge key"
                )
            if _finalize_calendar_key(
                mcal,
                f"Prepared morning calendar merge key from '{morning_study_date_column}'",
            ):
                return

        # --- Path 2: skyn slot order (row order = study day 1..N within burst) ---
        if mm is not None and 'metadata_skyn_first_day' in mm.columns:
            meta_first = mm[['SubID', 'Dataset_ID', 'metadata_skyn_first_day']].rename(
                columns={'metadata_skyn_first_day': '_meta_first'}
            )
            mcal = morning_sorted.merge(meta_first, on=['SubID', 'Dataset_ID'], how='left')
            unmatched_m = mcal['_meta_first'].isna()
            if unmatched_m.any():
                print(
                    f"  Warning: {int(unmatched_m.sum())} morning rows have no skyn row "
                    f"for (SubID, burst); dropping for slot merge"
                )
                mcal = mcal.loc[~unmatched_m].copy()
            if not mcal.empty:
                mcal = mcal.sort_values(['SubID', 'Dataset_ID', '_file_order'])
                mcal['morning_slot'] = mcal.groupby(['SubID', 'Dataset_ID'], sort=False).cumcount() + 1
                anchor = mcal['_meta_first'] + pd.to_timedelta(int(first_date_adjuster), unit='D')
                mcal['morning_merge_date'] = anchor + pd.to_timedelta(
                    mcal['morning_slot'] - 1, unit='D'
                )
                mcal['morning_merge_date'] = pd.to_datetime(
                    mcal['morning_merge_date'], errors='coerce'
                ).dt.normalize()
                mcal = mcal.drop(columns=['_meta_first'], errors='ignore')
                mcal_export = mcal.copy()
                mcal_export = self._morning_export_add_window_qc(mcal_export, first_date_adjuster)
                _sdt = pd.to_datetime(mcal_export['morning_merge_date'], errors='coerce')
                mcal_export['morning_calendar_date'] = _sdt.dt.strftime('%Y-%m-%d')
                mcal_export.loc[_sdt.isna(), 'morning_calendar_date'] = np.nan
                self.morning_annotated = mcal_export.drop(
                    columns=[c for c in mcal_export.columns if c.startswith('_')],
                    errors='ignore',
                )
                if _finalize_calendar_key(
                    mcal,
                    (
                        'Prepared morning slot merge (anchor = first_day + '
                        f'{first_date_adjuster} d; row order = study day 1..N)'
                    ),
                ):
                    return

        # --- Path 3: legacy submission order vs TAC day_no ---
        mleg = morning_sorted.copy()
        if response_type_col in mleg.columns:
            mleg = mleg[
                mleg[response_type_col].astype(str).str.strip() == submission_label
            ].copy()
        mleg['day_no'] = mleg.groupby(['SubID', 'Dataset_ID'], sort=False).cumcount() + 1
        mleg = mleg.drop_duplicates(subset=['SubID', 'Dataset_ID', 'day_no'], keep='last')
        mleg['morning_self_report_alcohol'] = self._morning_self_report_series(
            mleg, self_report_col, response_type_col, submission_label
        )
        attach_cols = self._prepare_morning_attach_columns(
            mleg, self._morning_extra_attach_cols
        )
        self.morning_merge_key_day_no = mleg[
            ['SubID', 'Dataset_ID', 'day_no', 'morning_self_report_alcohol'] + attach_cols
        ].copy()
        for c in ('SubID', 'Dataset_ID', 'day_no'):
            self.morning_merge_key_day_no[c] = pd.to_numeric(
                self.morning_merge_key_day_no[c], errors='coerce'
            ).astype(np.int64)
        print('  Prepared morning day_no merge key (Submissions only, k-th = day_no k).')
        if attach_cols:
            print(f"  Attach QC morning columns on day_no key: {', '.join(attach_cols)}")

    def add_morning_report_drink_agreement(
        self,
        morning_csv_path,
        pred_col='pred_drinking_day',
        self_report_col='mr_al_y',
        trigger_col='Trigger Name',
        id_col='id',
        response_type_col='Response Type',
        submission_label='Submission',
        morning_study_date_column=None,
        skyn_dates_csv_path=None,
        metadata_id_column='ID',
        metadata_burst_column='burst_id',
        first_day_column='first_day',
        last_date_column='last_day',
        first_date_adjuster=-1,
        extra_attach_cols=None,
    ):
        """
        Merge morning self-report onto ``self.day_features`` and add ``self_report_and_tac_comparison``
        (agreement vs ``pred_col``, default ``pred_drinking_day``; non-drinking days need ``prediction_reliable`` = 1).

        Builds merge keys via ``prepare_morning_self_report_for_tac_merge``, which uses
        ``self.skyn_dates_metadata`` from ``filter_days_by_date_range`` when set (else reads
        ``skyn_dates_csv_path``). Calendar join matches **SubID, Dataset_ID, begin_day date**;
        otherwise uses **day_no**. ``inside_burst`` is unchanged (set only in ``filter_days_by_date_range``).

        Optional ``extra_attach_cols`` are merged as QC-only morning fields (default: none).
        Cohort scripts should pass their column list explicitly.
        """
        attach_cols = list(extra_attach_cols or ())
        if not os.path.isfile(morning_csv_path):
            print(f"Warning: Morning report file not found: {morning_csv_path}")
            self.day_features['morning_self_report_alcohol'] = np.nan
            self.day_features['morning_report_matched'] = 0
            for c in attach_cols:
                self.day_features[c] = np.nan
            if pred_col in self.day_features.columns:
                self.day_features['self_report_and_tac_comparison'] = (
                    self._compute_self_report_and_tac_comparison(
                        self.day_features,
                        pred_col=pred_col,
                        sr_col='morning_self_report_alcohol',
                    )
                )
            else:
                self.day_features['self_report_and_tac_comparison'] = np.nan
            return

        if pred_col not in self.day_features.columns:
            print(f"Warning: '{pred_col}' missing — cannot label morning vs Skyn agreement")
            self.day_features['morning_self_report_alcohol'] = np.nan
            self.day_features['morning_report_matched'] = 0
            self.day_features['self_report_and_tac_comparison'] = np.nan
            for c in attach_cols:
                self.day_features[c] = np.nan
            return

        drop_cols = (
            'morning_self_report_alcohol',
            'morning_report_matched',
            'self_report_and_tac_comparison',
            'morning_merge_date',
            *attach_cols,
        )
        for col in drop_cols:
            if col in self.day_features.columns:
                self.day_features.drop(columns=[col], inplace=True)

        print("\nJoining morning self-report (mr_al_y) to day features...")

        self.prepare_morning_self_report_for_tac_merge(
            morning_csv_path=morning_csv_path,
            self_report_col=self_report_col,
            trigger_col=trigger_col,
            id_col=id_col,
            response_type_col=response_type_col,
            submission_label=submission_label,
            morning_study_date_column=morning_study_date_column,
            skyn_dates_csv_path=skyn_dates_csv_path,
            metadata_id_column=metadata_id_column,
            metadata_burst_column=metadata_burst_column,
            first_day_column=first_day_column,
            last_date_column=last_date_column,
            first_date_adjuster=first_date_adjuster,
            extra_attach_cols=attach_cols,
        )

        df = self.day_features.copy()

        # Expand to a full 14-day burst grid before merging, so SR can appear on TAC-missing days.
        meta = self._skyn_metadata_table_for_morning(
            skyn_dates_csv_path=skyn_dates_csv_path,
            metadata_id_column=metadata_id_column,
            metadata_burst_column=metadata_burst_column,
            first_day_column=first_day_column,
            last_date_column=last_date_column,
        )
        if meta is not None and len(meta):
            # Prefer an inferred day-start hour from observed begin_day; fallback to 6.
            _h = 6
            if 'begin_day' in df.columns:
                try:
                    _bdh = pd.to_datetime(df['begin_day'], errors='coerce').dt.hour.dropna()
                    if len(_bdh):
                        _h = int(_bdh.mode().iloc[0])
                except Exception:
                    _h = 6
            df = self._expand_to_full_14_day_burst_grid(
                df,
                skyn_meta=meta,
                n_days_per_burst=14,
                day_start_hour=_h,
            )
            print(
                "  Expanded to full 14-day burst grid using skyn dates metadata "
                f"(day_start_hour={_h}; includes TAC-missing days as synthetic rows)."
            )

        if self.morning_merge_key_calendar is not None and len(self.morning_merge_key_calendar):
            if 'begin_day' not in df.columns:
                print("Warning: day_features missing 'begin_day'; cannot calendar-merge morning")
                df['morning_self_report_alcohol'] = np.nan
                df['morning_merge_date'] = pd.NaT
                for c in attach_cols:
                    df[c] = np.nan
            else:
                df['SubID'] = pd.to_numeric(df['SubID'], errors='coerce').astype(np.int64)
                df['Dataset_ID'] = pd.to_numeric(df['Dataset_ID'], errors='coerce').astype(np.int64)
                df = self._merge_self_report_by_calendar_day(
                    df, self.morning_merge_key_calendar
                )
                print(
                    "  Merged morning self-report on SubID, Dataset_ID, calendar date (TAC begin_day)."
                )
        elif self.morning_merge_key_day_no is not None and len(self.morning_merge_key_day_no):
            merge_keys = ['SubID', 'Dataset_ID', 'day_no']
            missing = [c for c in merge_keys if c not in df.columns]
            if missing:
                print(f"Warning: day_features missing {missing} for morning day_no merge")
                df['morning_self_report_alcohol'] = np.nan
                df['morning_merge_date'] = pd.NaT
                for c in attach_cols:
                    df[c] = np.nan
            else:
                df['SubID'] = pd.to_numeric(df['SubID'], errors='coerce').astype(np.int64)
                df['Dataset_ID'] = pd.to_numeric(df['Dataset_ID'], errors='coerce').astype(np.int64)
                df['day_no'] = pd.to_numeric(df['day_no'], errors='coerce').astype(np.int64)
                df = df.merge(
                    self.morning_merge_key_day_no,
                    on=merge_keys,
                    how='left',
                )
                df['morning_merge_date'] = pd.NaT
                print("  Using legacy day_no merge (Submissions only).")
        else:
            print(
                "Warning: no morning merge key produced; check morning CSV, "
                "study-date column, or skyn metadata"
            )
            df['morning_self_report_alcohol'] = np.nan
            df['morning_merge_date'] = pd.NaT
            for c in attach_cols:
                df[c] = np.nan

        for c in attach_cols:
            if c not in df.columns:
                df[c] = np.nan

        if 'inside_burst' not in df.columns:
            df['inside_burst'] = np.nan

        matched = df['morning_self_report_alcohol'].notna()
        df['morning_report_matched'] = matched.astype(int)

        df['self_report_and_tac_comparison'] = self._compute_self_report_and_tac_comparison(
            df,
            pred_col=pred_col,
            sr_col='morning_self_report_alcohol',
        )

        # Unknown breakdown: assign a reason label for day rows that were not evaluable
        # (and for inserted scaffold rows). Blank for non-unknown rows that have observed TAC.
        unknown_keys = {'true_unknown', 'false_unknown', 'unknown_sr_missing'}
        is_unknown = df['self_report_and_tac_comparison'].astype(str).isin(unknown_keys)
        df['unknown_reason'] = ''
        if 'tac_day_missing' in df.columns:
            tac_missing = pd.to_numeric(df['tac_day_missing'], errors='coerce').fillna(0) == 1
        else:
            tac_missing = pd.Series(False, index=df.index)
        if 'prediction_reliable' in df.columns:
            unreliable = pd.to_numeric(df['prediction_reliable'], errors='coerce').fillna(0) != 1
        else:
            unreliable = pd.Series(True, index=df.index)

        # Always label inserted scaffold rows explicitly (these are "missing day (expanded)" in the Excel filter).
        df.loc[tac_missing, 'unknown_reason'] = 'missing_day_expanded'

        # Observed unknown days are non-drinking days with prediction_reliable != 1.
        df.loc[is_unknown & (~tac_missing) & unreliable, 'unknown_reason'] = 'prediction_unreliable'

        self.day_features = df

        print(f"  Rows with non-missing mr_al_y match: {int(self.day_features['morning_report_matched'].sum())}")
        if 'mr_numdk' in self.day_features.columns:
            n_numdk = int(pd.to_numeric(self.day_features['mr_numdk'], errors='coerce').notna().sum())
            print(f"  Rows with numeric mr_numdk: {n_numdk}")
        if attach_cols:
            print(f"  Attached morning QC columns: {', '.join(attach_cols)}")
        vq = self.day_features['self_report_and_tac_comparison'].value_counts()
        print(
            "  Morning vs TAC agreement (self_report_and_tac_comparison):\n"
            f"{vq.to_string()}"
        )

        self._append_morning_tac_summary_stats(pred_col=pred_col)

    def export_day_level_tac_and_self_report_workbook(
        self,
        output_path,
        morning_csv_path,
        plot_column='signal_processing_plot',
        pred_col='pred_drinking_day',
        row_interval=20,
        column_interval=12,
        x_scale=65 / 140,
        y_scale=90 / 182,
    ):
        """
        Export day-level TAC and self-report (morning EMA) validation grids to Excel.

        Sheets:
            - ``Methodology``: column names and how key fields are defined (drinking day, quality, self-report,
              agreement); matches ``pred_col`` passed to this export
            - ``Morning``: ``morning_annotated`` when available (slots/dates + optional
              ``morning_implied_in_skyn_window`` QC), else raw CSV
            - ``Morning_day_merge``: current ``day_features`` (includes merged morning columns and
              ``inside_burst`` from ``filter_days_by_date_range`` when that ran)
            - ``Variable_key``: short definitions for morning agreement columns, then a second table of
              selected day-level columns from ``report_guide.day_feature_descriptions`` (includes ``pred_col``)
            - ``Agreement_summary``: when ``inside_burst`` exists, counts of Skyn day rows **included**
              (1) vs **excluded** (0) for the burst/wear window; day-level self-report and TAC totals;
              agreement only (``self_report_and_tac_comparison``), including
              ``true_unknown`` / ``false_unknown``, ``unknown_sr_missing`` / ``unknown [SR missing]`` (no drinking day,
              ``prediction_reliable`` = 0, no morning), and no-morning rows split by drinking-day vs reliable non-drinking, plus a subtotal
              row, on the same ``inside_burst == 1`` subset as the burst grids (or all rows if ``inside_burst`` absent)
            - ``Burst_1``, ``Burst_2``, ``Burst_3``: **only** day rows with ``inside_burst == 1``;
              grid with days 1–14 on rows (TAC plot per cell); participant columns run left to right
              from fewer to more drinking days (count of ``pred_col`` = 1 days in that burst; ties keep
              groupby order). Per-participant agreement mini-tables use the same subset.
              One row above each plot shows a short agreement label (see ``agreement_burst_cell_display_label``;
              includes ``unknown [SR missing]`` for ``unknown_sr_missing``: non-drinking day, no morning SR, ``prediction_reliable`` = 0).

        Args:
            output_path: Path to ``.xlsx`` output
            morning_csv_path: Path to morning export CSV
            plot_column: Day-level filepath column to embed (default signal-processing plot)
            pred_col: TAC drinking-day flag (``pred_col`` = 1) for ranking and positivity
            row_interval / column_interval / x_scale / y_scale: layout matches ``embed_graphs`` defaults
        """
        hq_agreement_order = list(dayFeatures.HQ_AGREEMENT_ORDER_AND_LABELS)

        out_dir = os.path.dirname(output_path)
        if out_dir:
            os.makedirs(out_dir, exist_ok=True)

        print(f"\nExporting day-level TAC + self-report workbook: {output_path}")

        with pd.ExcelWriter(output_path, engine='xlsxwriter', mode='w') as writer:
            book = writer.book
            bold = book.add_format({'bold': True})
            wrap_text = book.add_format({'text_wrap': True, 'valign': 'top'})

            meth_df = self.build_tac_self_report_methodology_dataframe(pred_col=pred_col)
            meth_df.to_excel(writer, sheet_name='Methodology', index=False)
            ws_meth = writer.sheets['Methodology']
            ws_meth.set_column(0, 0, 38, wrap_text)
            ws_meth.set_column(1, 1, 108, wrap_text)
            ws_meth.freeze_panes(1, 0)

            # --- Morning (annotated: slot/calendar + optional QC vs skyn window) ---
            if getattr(self, 'morning_annotated', None) is not None and len(self.morning_annotated):
                self.morning_annotated.to_excel(writer, sheet_name='Morning', index=False)
            elif morning_csv_path and os.path.isfile(morning_csv_path):
                morning_raw = pd.read_csv(morning_csv_path)
                morning_raw.to_excel(writer, sheet_name='Morning', index=False)
            else:
                pd.DataFrame({'note': ['morning.csv not found']}).to_excel(
                    writer, sheet_name='Morning', index=False
                )

            # --- Merged day-level (+ morning columns) ---
            self.day_features.to_excel(writer, sheet_name='Morning_day_merge', index=False)

            vk_df = self.build_tac_self_report_variable_key_dataframe(pred_col=pred_col)
            vk_df.to_excel(writer, sheet_name='Variable_key', index=False)
            ws_vk = writer.sheets['Variable_key']
            ws_vk.set_column(0, 0, 40, wrap_text)
            ws_vk.set_column(1, 1, 92, wrap_text)
            ws_vk.freeze_panes(1, 0)

            df_all = self.day_features.copy()
            df_all['SubID'] = pd.to_numeric(df_all['SubID'], errors='coerce')
            df_all['Dataset_ID'] = pd.to_numeric(df_all['Dataset_ID'], errors='coerce')
            df_all['day_no'] = pd.to_numeric(df_all['day_no'], errors='coerce')

            if 'inside_burst' in df_all.columns:
                n_before = len(df_all)
                df_all = df_all[df_all['inside_burst'] == 1].copy()
                n_drop = n_before - len(df_all)
                if n_drop:
                    print(
                        f"  Morning/TAC burst grids: using {len(df_all)} day rows "
                        f"(inside_burst==1); excluded {n_drop}"
                    )
            else:
                print(
                    "  Warning: 'inside_burst' missing — burst grids use all day rows "
                    "(run filter_days_by_date_range to attach TAC-date vs skyn window flag)"
                )

            summary_df = self._morning_tac_summary_blocks(
                df_all,
                pred_col=pred_col,
                hq_agreement_col='self_report_and_tac_comparison',
                df_full_burst_counts=self.day_features,
            )
            summary_df.to_excel(writer, sheet_name='Agreement_summary', index=False, startrow=1)
            ws_sum = writer.sheets['Agreement_summary']

            # Agreement label grid (3x3) as its own table at top-right.
            grid_df = self._agreement_label_grid_results(
                df_all,
                agreement_col='self_report_and_tac_comparison',
            )
            if grid_df is not None and len(grid_df):
                # MultiIndex-style layout: "Self report" over SR columns; "TAC" beside row labels.
                grid_hdr = book.add_format(
                    {'bold': True, 'align': 'center', 'valign': 'vcenter', 'text_wrap': True}
                )
                grid_cell = book.add_format({'valign': 'vcenter', 'text_wrap': True})
                tac_corner_col = 3
                row_label_col = 4
                data_col0 = 5
                sr_cols = [c for c in grid_df.columns if c != 'TAC \\ self-report']
                last_data_col = data_col0 + len(sr_cols) - 1
                r_title = 1
                r_sr_outer = 2
                r_sr_sub = 3
                r_data0 = 4
                n_rows = len(grid_df)

                ws_sum.merge_range(
                    r_title,
                    tac_corner_col,
                    r_title,
                    last_data_col,
                    '— Agreement label grid (n and % of total days) —',
                    grid_hdr,
                )
                ws_sum.merge_range(
                    r_sr_outer,
                    tac_corner_col,
                    r_sr_outer,
                    row_label_col,
                    '',
                    grid_hdr,
                )
                ws_sum.merge_range(
                    r_sr_outer,
                    data_col0,
                    r_sr_outer,
                    last_data_col,
                    'Self Report',
                    grid_hdr,
                )
                for j, col_name in enumerate(sr_cols):
                    ws_sum.write(r_sr_sub, data_col0 + j, col_name, bold)
                ws_sum.merge_range(
                    r_data0,
                    tac_corner_col,
                    r_data0 + n_rows - 1,
                    tac_corner_col,
                    'TAC',
                    grid_hdr,
                )
                label_col = 'TAC \\ self-report'
                for i in range(n_rows):
                    ws_sum.write(
                        r_data0 + i,
                        row_label_col,
                        grid_df[label_col].iloc[i],
                        grid_cell,
                    )
                    for j, col_name in enumerate(sr_cols):
                        ws_sum.write(
                            r_data0 + i,
                            data_col0 + j,
                            grid_df[col_name].iloc[i],
                            grid_cell,
                        )
                ws_sum.set_column(tac_corner_col, tac_corner_col, 10)
                ws_sum.set_column(row_label_col, row_label_col, 28)
                ws_sum.set_column(data_col0, last_data_col, 30)

            if plot_column not in df_all.columns:
                print(f"  Warning: plot column '{plot_column}' missing; burst sheets will show placeholders")

            for burst_id in (1, 2, 3):
                sheet_name = f'Burst_{burst_id}'
                lb = df_all[df_all['Dataset_ID'] == burst_id].copy()
                ws = book.add_worksheet(sheet_name)

                if lb.empty:
                    ws.write(0, 0, f'No day-level rows for burst {burst_id}')
                    continue

                drink_days_per_subid = (
                    lb.groupby('SubID', sort=False)[pred_col]
                    .apply(lambda s: (pd.to_numeric(s, errors='coerce').fillna(0) > 0).sum())
                    .sort_values(ascending=True)
                )
                subids = [int(x) for x in drink_days_per_subid.index.tolist()]

                # Agreement table: header row 2, one row per ``hq_agreement_order`` key; label row below; then plots.
                n_ag_rows = len(hq_agreement_order)
                first_plot_row = 2 + n_ag_rows + 2

                agree_lbl_fmt = book.add_format({'text_wrap': True, 'valign': 'bottom', 'font_size': 9})
                agree_mini_fmt = book.add_format({'text_wrap': True, 'valign': 'vcenter', 'font_size': 9})
                burst_ordering_fmt = book.add_format({'bold': True, 'text_wrap': True, 'valign': 'top'})
                hq_agreement_col = 'self_report_and_tac_comparison'

                ws.write(0, 0, f'Burst {burst_id}: TAC+ rank (inside_burst==1 only)', bold)
                ws.write(
                    1,
                    0,
                    'Ordering: From left to right, SubIDs are sorted from fewer to more drinking days.',
                    burst_ordering_fmt,
                )

                for j, subid in enumerate(subids):
                    col_start = 2 + j * column_interval
                    ws.write(0, col_start, f'SubID {subid}', bold)
                    ws.write(1, col_start, f'TAC+ days: {int(drink_days_per_subid.loc[subid])}', bold)

                    sub_b = lb[lb['SubID'] == subid]
                    if hq_agreement_col in sub_b.columns:
                        vc = sub_b[hq_agreement_col].astype(str).value_counts()
                    else:
                        vc = pd.Series(dtype=int)
                    denom_p = max(1, len(sub_b))

                    ws.merge_range(2, col_start, 2, col_start + 1, 'Agreement', bold)
                    r = 3
                    for key, label in hq_agreement_order:
                        n_k = int(vc.get(key, 0))
                        pct_p = 100.0 * n_k / denom_p
                        ws.merge_range(
                            r,
                            col_start,
                            r,
                            col_start + 1,
                            f'{label}: {n_k} ({pct_p:.1f}%)',
                            agree_mini_fmt,
                        )
                        r += 1

                for d in range(1, 15):
                    r = first_plot_row + (d - 1) * row_interval
                    ws.write(r, 0, f'Day {d}', bold)

                    for j, subid in enumerate(subids):
                        col_start = 2 + j * column_interval
                        rows_d = lb[(lb['SubID'] == subid) & (lb['day_no'] == d)]
                        if rows_d.empty:
                            ws.write(r, col_start, '—')
                            continue
                        hq_v = (
                            rows_d[hq_agreement_col].iloc[0]
                            if hq_agreement_col in rows_d.columns
                            else np.nan
                        )
                        cell_lbl = self.agreement_burst_cell_display_label(hq_v)
                        ws.write(r - 1, col_start, cell_lbl, agree_lbl_fmt)

                        if plot_column not in rows_d.columns:
                            ws.write(r, col_start, 'no plot column')
                            continue
                        raw_p = rows_d[plot_column].iloc[0]
                        if raw_p is None or (isinstance(raw_p, float) and np.isnan(raw_p)):
                            ws.write(r, col_start, 'no plot')
                            continue
                        p = str(raw_p).strip()
                        if not p or p.lower() == 'nan':
                            ws.write(r, col_start, 'no plot')
                            continue
                        if os.path.isfile(p):
                            try:
                                ws.insert_image(
                                    r,
                                    col_start,
                                    p,
                                    {'x_scale': x_scale, 'y_scale': y_scale},
                                )
                            except Exception:
                                ws.write(r, col_start, f'Invalid: {p[:40]}...')
                        else:
                            ws.write(r, col_start, f'Missing file: {p[:50]}')

                ws.set_column(0, 0, 56)
                ws.set_column(1, 1, 3)

        print('  Done.')

    _CURVE_ATTACH_PLOT_COLS = frozenset(
        {
            'smoothed_curve_plot',
            'signal_processing_plot',
            'device_removal_plot',
            'signal_processing_plot_wide',
        }
    )

    @staticmethod
    def _curve_overlap_slot_indices(columns):
        pat = re.compile(r'^curve_(\d+)_id$')
        slots = []
        for c in columns:
            m = pat.match(str(c))
            if m:
                slots.append(int(m.group(1)))
        return sorted(set(slots))

    @staticmethod
    def _curve_overlap_slot_gated_out_for_hq_selection(row: pd.Series, n: int) -> bool:
        """
        Exclude from max-HQ ranking when the curve **started before** the social day and has a small on-day
        footprint: ``overlap_hours`` < :py:data:`_HQ_PICK_PRIOR_DAY_MIN_OVERLAP_HOURS` **or** (when
        ``curve_{n}_duration_CURVE`` is known and positive) overlap fraction of full curve duration <
        :py:data:`_HQ_PICK_PRIOR_DAY_MIN_CURVE_FRACTION_IN_DAY`.
        """
        ep = f'curve_{n}_extends_prior_day'
        oh = f'curve_{n}_overlap_hours'
        dur_c = f'curve_{n}_duration_CURVE'
        extends_prior = bool(pd.to_numeric(row[ep], errors='coerce') == 1) if ep in row.index else False
        if not extends_prior:
            return False
        overlap_h = pd.to_numeric(row[oh], errors='coerce') if oh in row.index else np.nan
        dur_h = pd.to_numeric(row[dur_c], errors='coerce') if dur_c in row.index else np.nan
        low_hours = (not pd.notna(overlap_h)) or (float(overlap_h) < _HQ_PICK_PRIOR_DAY_MIN_OVERLAP_HOURS)
        low_frac = False
        if pd.notna(dur_h) and float(dur_h) > 0 and pd.notna(overlap_h):
            low_frac = float(overlap_h) / float(dur_h) < _HQ_PICK_PRIOR_DAY_MIN_CURVE_FRACTION_IN_DAY
        return bool(low_frac or low_hours)

    @staticmethod
    def _pick_curve_overlap_longest_hq_row(
        row: pd.Series, slots_sorted: list[int]
    ) -> Optional[Tuple[int, int]]:
        """
        Among overlap slots with a non-null ``curve_{n}_id``, pick the same winner as curve attachment:
        max ``curve_{n}_high_quality_duration``, then lower ``n``, then lower ``curve_id``.

        Slots gated out by ``_curve_overlap_slot_gated_out_for_hq_selection`` (prior-day start with little curve
        on the social day) are dropped unless that would remove **all** candidates, in which case gating is ignored
        for that row.

        For curves built in ``Curve``, ``high_quality_duration_CURVE`` is always a finite float (see
        ``DataQualityAnalyzer.get_high_quality_duration``); day overlap rows use ``.get(..., 0)``. So in normal
        exports every candidate has a non-NaN HQ value (possibly 0). The branch that sorts only by ``(n,
        curve_id)`` applies only when **every** HQ coerces to NaN (e.g. truncated CSV / missing columns) and is
        defensive. Returns ``(slot_n, curve_id_int)`` or ``None``.
        """
        cand = []
        for n in slots_sorted:
            ic = f'curve_{n}_id'
            hc = f'curve_{n}_high_quality_duration'
            if ic not in row.index:
                continue
            cid = row[ic]
            if pd.isna(cid):
                continue
            try:
                cid_i = int(cid)
            except (ValueError, TypeError):
                continue
            hq = pd.to_numeric(row[hc], errors='coerce') if hc in row.index else np.nan
            cand.append((n, cid_i, hq))
        if not cand:
            return None
        eligible = [t for t in cand if not dayFeatures._curve_overlap_slot_gated_out_for_hq_selection(row, t[0])]
        if eligible:
            cand = eligible
        has_hq = any(pd.notna(c[2]) for c in cand)
        if has_hq:
            cand.sort(
                key=lambda t: (
                    -(float(t[2]) if pd.notna(t[2]) else float('-inf')),
                    t[0],
                    t[1],
                )
            )
        else:
            cand.sort(key=lambda t: (t[0], t[1]))
        return cand[0][0], cand[0][1]

    @staticmethod
    def pick_curve_attach_id_longest_hq_overlap(out: pd.DataFrame, slots: list[int]) -> pd.Series:
        """
        Per day row, choose one ``curve_id`` among overlap slots ``curve_{n}_id`` that maximizes
        ``curve_{n}_high_quality_duration`` (already on ``out`` from overlap detection; sourced from
        ``high_quality_duration_CURVE`` in curve features). Prior-day-starting curves with
        ``overlap_hours`` < 1 or overlap / ``curve_{n}_duration_CURVE`` < 0.5 are omitted when other candidates exist.

        Tie-break: higher HQ duration wins; then smaller overlap slot index ``n``; then smaller ``curve_id``.

        If every overlap HQ value is NaN after coercion (not expected for pipeline-produced curves), falls back
        to earliest slot with a non-null ``curve_{n}_id``; see ``_pick_curve_overlap_longest_hq_row``.
        """
        slots_sorted = sorted(slots)

        def pick_row(row: pd.Series):
            t = dayFeatures._pick_curve_overlap_longest_hq_row(row, slots_sorted)
            return float(t[1]) if t else np.nan

        return out.apply(pick_row, axis=1)

    @staticmethod
    def pick_curve_overlap_slot_longest_hq_overlap(
        out: pd.DataFrame, slots: Optional[List[int]] = None
    ) -> pd.Series:
        """
        Same winner as ``pick_curve_attach_id_longest_hq_overlap``, but returns overlap slot index ``n``
        (float) instead of ``curve_id``. NaN when no overlapping curve in any slot.
        """
        slots_sorted = sorted(slots if slots is not None else dayFeatures._curve_overlap_slot_indices(out.columns))

        def pick_row(row: pd.Series):
            t = dayFeatures._pick_curve_overlap_longest_hq_row(row, slots_sorted)
            return float(t[0]) if t else np.nan

        return out.apply(pick_row, axis=1)

    @staticmethod
    def _resolve_curve_columns_to_attach(curve_features_df, curve_feature_columns):
        from App.SDM.Documenting.report_guide import report_guide

        skip = {'subid', 'dataset_id', 'curve_id', 'dataset_identifier'}
        if curve_feature_columns is not None:
            out = []
            for c in curve_feature_columns:
                if c in curve_features_df.columns and c not in skip and c not in dayFeatures._CURVE_ATTACH_PLOT_COLS:
                    out.append(c)
            return out

        stats = list(report_guide.stats_features)
        # Periphery + extra curve quality metrics for max-HQ attachment (ML / annotation workbook).
        _periphery_quality_stems = (
            'total_duration',
            'device_turned_on_percent',
            'device_worn_percent',
            'imputed_percent',
            'total_low_quality_percent',
            'unimputed_low_quality_percent',
            'total_gap_percent',
            'total_non_wear_percent',
            'total_jump_percent',
            'total_plummet_percent',
            'total_extreme_negative_percent',
            'low_quality_imputation_ratio',
        )
        _periphery_extras = [
            f'{stem}_PERIPHERY_BEFORE' for stem in _periphery_quality_stems
        ] + [f'{stem}_PERIPHERY_AFTER' for stem in _periphery_quality_stems]
        _curve_quality_extras = [
            'unimputed_low_quality_percent_CURVE',
            'total_gap_percent_CURVE',
            'total_non_wear_percent_CURVE',
            'total_jump_percent_CURVE',
            'total_plummet_percent_CURVE',
            'total_extreme_negative_percent_CURVE',
            'low_quality_imputation_ratio_CURVE',
        ]
        extras = [
            'DRINKING_PRED',
            'REGION_VALID',
            'CURVE_VALID',
            'begin_CURVE',
            'end_CURVE',
            'curve_threshold',
            'required_HQ_duration',
            'high_quality_duration_CURVE',
            'high_quality_percent_CURVE',
            'duration_CURVE',
            'peak_CURVE',
            'auc_total_CURVE',
            'auc_relative_CURVE',
            'rise_rate_CURVE',
            'fall_rate_CURVE',
            'rise_rate_point_to_point_CURVE',
            'fall_rate_point_to_point_CURVE',
            'rise_duration_CURVE',
            'fall_duration_CURVE',
            'FLAG_below_threshold_curve',
            'FLAG_incomplete_curve_start_curve',
        ] + _periphery_extras + _curve_quality_extras
        wanted = []
        for c in stats + extras:
            if (
                c in curve_features_df.columns
                and c not in skip
                and c not in dayFeatures._CURVE_ATTACH_PLOT_COLS
                and c not in wanted
            ):
                wanted.append(c)
        return wanted

    # ``top_hq_curve_<metric>`` day columns -> median of that metric over prior study days in the same burst
    # (build_day_features_with_curve_attachments; sources are max-HQ overlap attach columns).
    _PRIOR_MEDIAN_CURVE_METRICS = (
        'peak_CURVE',
        'auc_total_CURVE',
        'duration_CURVE',
        'high_quality_duration_CURVE',
        'rise_rate_CURVE',
        'fall_rate_CURVE',
        'rise_rate_point_to_point_CURVE',
        'fall_rate_point_to_point_CURVE',
        'fall_duration_CURVE',
        'rise_duration_CURVE',
    )
    _PRIOR_MEDIAN_CURVE_FEATURE_SOURCES = tuple(
        (_TOP_HQ_CURVE_COL_PREFIX + m, 'median_prior_' + m) for m in _PRIOR_MEDIAN_CURVE_METRICS
    )

    # ``curve_{n}_{suffix}`` on overlap slots -> ``hq_best_*`` for the max-HQ slot (same winner as attachment).
    _HQ_BEST_OVERLAP_SOURCES = (
        ('id', 'hq_best_curve_id'),
        ('overlap_hours', 'hq_best_overlap_hours'),
        ('high_quality_duration', 'hq_best_high_quality_duration'),
        ('extends_prior_day', 'hq_best_extends_prior_day'),
        ('extends_next_day', 'hq_best_extends_next_day'),
    )

    @staticmethod
    def assign_hq_best_curve_overlap_columns(df: pd.DataFrame) -> None:
        """
        Add ``hq_best_*`` columns by copying overlap-slot fields from the ``curve_{n}_*`` slot with the largest
        ``curve_{n}_high_quality_duration`` (same tie-break as ``pick_curve_attach_id_longest_hq_overlap``).

        Aligns day-level overlap descriptors with the **max-HQ** curve, distinct from time-first ``curve_1_*``.
        Mutates ``df`` in place.
        """
        slots = dayFeatures._curve_overlap_slot_indices(df.columns)
        slot_order = sorted(slots)
        df['hq_best_curve_slot'] = np.nan
        if not slot_order:
            for _, dst in dayFeatures._HQ_BEST_OVERLAP_SOURCES:
                df[dst] = np.nan
            return

        win = dayFeatures.pick_curve_overlap_slot_longest_hq_overlap(df, slot_order)
        df['hq_best_curve_slot'] = win
        idx_map = {n: j for j, n in enumerate(slot_order)}
        win_arr = pd.to_numeric(win, errors='coerce').to_numpy(dtype=float, copy=False)

        for src_suffix, dst in dayFeatures._HQ_BEST_OVERLAP_SOURCES:
            mat = np.full((len(df), len(slot_order)), np.nan, dtype=float)
            for j, n in enumerate(slot_order):
                c = f'curve_{n}_{src_suffix}'
                if c in df.columns:
                    mat[:, j] = pd.to_numeric(df[c], errors='coerce').to_numpy(dtype=float, copy=False)
            picked = np.full(len(df), np.nan, dtype=float)
            for i in range(len(df)):
                ws = win_arr[i]
                if np.isnan(ws):
                    continue
                j = idx_map.get(int(ws))
                if j is None:
                    continue
                picked[i] = mat[i, j]
            df[dst] = picked

    @staticmethod
    def assign_hq_best_overlap_fraction_total_duration(df: pd.DataFrame) -> None:
        """
        ``hq_best_overlap_hours`` / ``top_hq_curve_total_duration_CURVE`` when denominator > 0; else NaN.
        Requires ``assign_hq_best_curve_overlap_columns`` and merged ``top_hq_curve_total_duration_CURVE`` (or NaN).
        """
        col_out = 'hq_best_overlap_fraction_total_duration_CURVE'
        ov = 'hq_best_overlap_hours'
        dur = f'{_TOP_HQ_CURVE_COL_PREFIX}total_duration_CURVE'
        if ov not in df.columns:
            df[col_out] = np.nan
            return
        oh = pd.to_numeric(df[ov], errors='coerce')
        if dur not in df.columns:
            df[col_out] = np.nan
            return
        d = pd.to_numeric(df[dur], errors='coerce')
        with np.errstate(divide='ignore', invalid='ignore'):
            frac = oh / d
        ok = (d > 0) & np.isfinite(d) & oh.notna()
        df[col_out] = frac.where(ok)

    @staticmethod
    def merge_top_hq_curve_feature_columns_into_day_df(
        out: pd.DataFrame,
        curve_features_df: pd.DataFrame,
        curve_feature_columns=None,
    ) -> None:
        """
        Left-merge curve-level columns onto ``out`` as ``top_hq_curve_*`` using max-HQ ``curve_id`` per row.
        Mutates ``out`` in place. No-op if ``attach_cols`` resolves empty.
        """
        slots = dayFeatures._curve_overlap_slot_indices(out.columns)
        attach_cols = dayFeatures._resolve_curve_columns_to_attach(curve_features_df, curve_feature_columns)
        attach_cols = [c for c in attach_cols if c in curve_features_df.columns]
        if not attach_cols:
            print(
                '  Warning: no curve columns selected for attachment (check curve_feature_columns / overlap with curve_features_df)'
            )
            return

        cf = curve_features_df.copy()
        cf['_subid'] = pd.to_numeric(cf['subid'], errors='coerce')
        cf['_dataset_id'] = pd.to_numeric(cf['dataset_id'], errors='coerce')
        cf['_curve_id'] = pd.to_numeric(cf['curve_id'], errors='coerce')
        cf = cf.dropna(subset=['_subid', '_dataset_id', '_curve_id'])
        cf = cf.drop_duplicates(subset=['_subid', '_dataset_id', '_curve_id'], keep='first')

        cc = [c for c in attach_cols if c in cf.columns]
        if not cc:
            return

        attach_curve_ids = dayFeatures.pick_curve_attach_id_longest_hq_overlap(out, slots)
        print(
            '  Curve attachment: merged curve stats under top_hq_curve_* from overlap curve with max '
            'high_quality_duration_CURVE (uses day columns curve_{n}_high_quality_duration).'
        )
        left = pd.DataFrame(
            {
                '_row_ix': np.asarray(out.index),
                'SubID': pd.to_numeric(out['SubID'], errors='coerce'),
                'Dataset_ID': pd.to_numeric(out['Dataset_ID'], errors='coerce'),
                '_merge_cid': pd.to_numeric(attach_curve_ids, errors='coerce'),
            }
        )
        left = left.dropna(subset=['SubID', 'Dataset_ID', '_merge_cid'])
        left['subid'] = left['SubID'].astype(np.int64)
        left['dataset_id'] = left['Dataset_ID'].astype(np.int64)
        left['curve_id'] = left['_merge_cid'].astype(np.int64)

        right = cf[['_subid', '_dataset_id', '_curve_id'] + cc].copy()
        right = right.rename(
            columns={
                '_subid': 'subid',
                '_dataset_id': 'dataset_id',
                '_curve_id': 'curve_id',
            }
        )
        right = right.drop_duplicates(subset=['subid', 'dataset_id', 'curve_id'], keep='first')

        j = left.merge(right, on=['subid', 'dataset_id', 'curve_id'], how='left')
        thq = dayFeatures.TOP_HQ_CURVE_COL_PREFIX
        prefixed = [f'{thq}{c}' for c in cc]
        rename_map = {c: f'{thq}{c}' for c in cc}
        j = j.rename(columns=rename_map)
        for col in prefixed:
            if col not in out.columns:
                out[col] = np.nan
        ix = j['_row_ix'].astype(int).to_numpy()
        out.loc[ix, prefixed] = j[prefixed].to_numpy()

    @staticmethod
    def enrich_day_features_top_hq_merge_hq_best_fraction_medians(
        df: pd.DataFrame,
        curve_features_df: pd.DataFrame,
        curve_feature_columns=None,
    ) -> None:
        """
        Standard post-overlap enrichment: merge ``top_hq_curve_*``, ``hq_best_*``,
        ``hq_best_overlap_fraction_total_duration_CURVE``, and ``median_prior_*`` (burst prior-day medians).
        """
        if curve_features_df is None or getattr(curve_features_df, 'empty', True):
            dayFeatures.assign_hq_best_curve_overlap_columns(df)
            dayFeatures.assign_hq_best_overlap_fraction_total_duration(df)
            dayFeatures.assign_median_prior_merged_curve_features_within_burst(df)
            return
        dayFeatures.merge_top_hq_curve_feature_columns_into_day_df(df, curve_features_df, curve_feature_columns)
        dayFeatures.assign_hq_best_curve_overlap_columns(df)
        dayFeatures.assign_hq_best_overlap_fraction_total_duration(df)
        dayFeatures.assign_median_prior_merged_curve_features_within_burst(df)

    @staticmethod
    def assign_median_prior_merged_curve_features_within_burst(df: pd.DataFrame) -> None:
        """
        For each row, add columns ``median_prior_*`` = median of the same merged ``top_hq_curve_*`` metric over
        **earlier** ``day_no`` within ``SubID`` and ``Dataset_ID``. First study day per burst: NaN.

        A prior day contributes only if ``top_hq_curve_high_quality_duration_CURVE`` (or
        ``hq_best_high_quality_duration`` if the merged column is absent) is ≥
        ``_PRIOR_MEDIAN_MIN_HIGH_QUALITY_DURATION_HOURS`` (1 hour).

        Uses the max-HQ merged attach columns (e.g. ``top_hq_curve_peak_CURVE``). Missing source columns are skipped.
        """
        req = ('SubID', 'Dataset_ID', 'day_no')
        if not all(c in df.columns for c in req):
            print(f'  Warning: missing columns {req}; skipping prior median curve features')
            return
        for _, dst in dayFeatures._PRIOR_MEDIAN_CURVE_FEATURE_SOURCES:
            df[dst] = np.nan
        active = [(src, dst) for src, dst in dayFeatures._PRIOR_MEDIAN_CURVE_FEATURE_SOURCES if src in df.columns]
        missing_src = [src for src, _ in dayFeatures._PRIOR_MEDIAN_CURVE_FEATURE_SOURCES if src not in df.columns]
        if missing_src:
            print(
                '  Note: prior median curve features omitted for missing merged columns: '
                f'{missing_src[:8]}{"..." if len(missing_src) > 8 else ""}'
            )
        if not active:
            return
        for _, g in df.groupby(['SubID', 'Dataset_ID'], sort=False):
            dn = pd.to_numeric(g['day_no'], errors='coerce')
            order = np.argsort(dn.to_numpy(), kind='mergesort')
            g_sorted = g.iloc[order]
            g_ix = g_sorted.index.to_numpy()
            n = len(g_ix)
            collected = {dst: [] for _, dst in active}
            out_blocks = {dst: np.empty(n, dtype=float) for _, dst in active}
            hq_merged = f'{_TOP_HQ_CURVE_COL_PREFIX}high_quality_duration_CURVE'
            for pos in range(n):
                ix = g_ix[pos]
                for src, dst in active:
                    out_blocks[dst][pos] = (
                        float(np.median(collected[dst])) if collected[dst] else np.nan
                    )
                if hq_merged in df.columns:
                    hq_v = pd.to_numeric(df.loc[ix, hq_merged], errors='coerce')
                elif 'hq_best_high_quality_duration' in df.columns:
                    hq_v = pd.to_numeric(df.loc[ix, 'hq_best_high_quality_duration'], errors='coerce')
                else:
                    hq_v = np.nan
                include_prior = pd.notna(hq_v) and float(hq_v) >= _PRIOR_MEDIAN_MIN_HIGH_QUALITY_DURATION_HOURS
                if include_prior:
                    for src, dst in active:
                        v = pd.to_numeric(df.loc[ix, src], errors='coerce')
                        if pd.notna(v):
                            collected[dst].append(float(v))
            for _, dst in active:
                ser_ord = pd.Series(out_blocks[dst], index=g_ix)
                df.loc[g.index, dst] = ser_ord.reindex(g.index)

    @staticmethod
    def assign_median_prior_from_all_qualifying_curves(df: pd.DataFrame) -> None:
        """
        For each row, add ``median_prior_*`` = median of the metric across **all** qualifying
        curves on **earlier** ``day_no`` within the same ``SubID`` / ``Dataset_ID`` burst.

        A curve in slot ``n`` qualifies if:
          - ``curve_{n}_high_quality_percent_CURVE`` >= ``_MEDIAN_PRIOR_MIN_HQ_PERCENT`` (0.75)
          - ``curve_{n}_high_quality_duration`` (hours) >= ``_MEDIAN_PRIOR_MIN_HQ_DURATION_HOURS`` (1.0)

        First day in a burst always gets NaN.  If no qualifying prior curves exist, all
        ``median_prior_*`` columns are NaN for that row.
        """
        req = ('SubID', 'Dataset_ID', 'day_no')
        if not all(c in df.columns for c in req):
            print(f'  Warning: missing columns {req}; skipping prior median curve features')
            return

        metrics = _PRIOR_MEDIAN_CURVE_METRICS_V2
        dst_names = [f'median_prior_{m}' for m in metrics]
        for dst in dst_names:
            df[dst] = np.nan

        slots = list(range(1, _MAX_CURVE_SLOTS + 1))
        slot_src_cols: list[list[tuple[str, str]]] = []
        for m, dst in zip(metrics, dst_names):
            pairs = []
            for n in slots:
                src = f'curve_{n}_{m}'
                if src in df.columns:
                    pairs.append((src, dst))
            slot_src_cols.append(pairs)

        hq_pct_cols = [f'curve_{n}_high_quality_percent_CURVE' for n in slots]
        hq_dur_cols = [f'curve_{n}_high_quality_duration' for n in slots]

        for _, g in df.groupby(['SubID', 'Dataset_ID'], sort=False):
            dn = pd.to_numeric(g['day_no'], errors='coerce')
            order = np.argsort(dn.to_numpy(), kind='mergesort')
            g_sorted = g.iloc[order]
            g_ix = g_sorted.index.to_numpy()
            n_rows = len(g_ix)

            collected: dict[str, list[float]] = {dst: [] for dst in dst_names}
            out_blocks: dict[str, np.ndarray] = {dst: np.empty(n_rows, dtype=float) for dst in dst_names}

            for pos in range(n_rows):
                ix = g_ix[pos]
                for dst in dst_names:
                    out_blocks[dst][pos] = float(np.median(collected[dst])) if collected[dst] else np.nan

                for n in slots:
                    hq_pct_col = hq_pct_cols[n - 1]
                    hq_dur_col = hq_dur_cols[n - 1]
                    if hq_pct_col not in df.columns or hq_dur_col not in df.columns:
                        continue
                    cid_col = f'curve_{n}_id'
                    if cid_col in df.columns and pd.isna(df.loc[ix, cid_col]):
                        continue
                    hq_pct = pd.to_numeric(df.loc[ix, hq_pct_col], errors='coerce')
                    hq_dur = pd.to_numeric(df.loc[ix, hq_dur_col], errors='coerce')
                    if pd.notna(hq_pct) and hq_pct > 1.0:
                        hq_pct_norm = hq_pct / 100.0
                    else:
                        hq_pct_norm = hq_pct
                    if not (pd.notna(hq_pct_norm) and hq_pct_norm >= _MEDIAN_PRIOR_MIN_HQ_PERCENT
                            and pd.notna(hq_dur) and hq_dur >= _MEDIAN_PRIOR_MIN_HQ_DURATION_HOURS):
                        continue
                    for pairs in slot_src_cols:
                        for src, dst in pairs:
                            if src.startswith(f'curve_{n}_'):
                                v = pd.to_numeric(df.loc[ix, src], errors='coerce')
                                if pd.notna(v):
                                    collected[dst].append(float(v))

            for dst in dst_names:
                ser_ord = pd.Series(out_blocks[dst], index=g_ix)
                df.loc[g.index, dst] = ser_ord.reindex(g.index)

    _MEDIAN_PRIOR_DAY_LEVEL_COLS = (
        'tac_q1_mean', 'tac_q1_max', 'tac_q1_auc',
        'tac_q2_q4_mean', 'tac_q2_q4_max', 'tac_q2_q4_auc',
        'curve_1_start_offset_hours', 'curve_1_end_offset_hours',
    )

    @staticmethod
    def assign_median_prior_day_region_tac(df: pd.DataFrame) -> None:
        """
        For each row, add ``median_prior_tac_<region>_<stat>`` = expanding median
        of the day-level TAC region feature across **earlier** ``day_no`` within
        the same ``SubID`` / ``Dataset_ID`` burst.

        Source columns: ``tac_q1_mean``, ``tac_q1_max``, ``tac_q1_auc``,
        ``tac_q2_q4_mean``, ``tac_q2_q4_max``, ``tac_q2_q4_auc``.
        First day in a burst always gets NaN.
        """
        req = ('SubID', 'Dataset_ID', 'day_no')
        if not all(c in df.columns for c in req):
            return

        src_cols = [c for c in dayFeatures._MEDIAN_PRIOR_DAY_LEVEL_COLS if c in df.columns]
        if not src_cols:
            return

        dst_map = {src: f'median_prior_{src}' for src in src_cols}
        for dst in dst_map.values():
            df[dst] = np.nan

        for _, g in df.groupby(['SubID', 'Dataset_ID'], sort=False):
            dn = pd.to_numeric(g['day_no'], errors='coerce')
            order = np.argsort(dn.to_numpy(), kind='mergesort')
            g_sorted = g.iloc[order]
            g_ix = g_sorted.index.to_numpy()
            n_rows = len(g_ix)

            collected: dict[str, list[float]] = {dst: [] for dst in dst_map.values()}
            out_blocks: dict[str, np.ndarray] = {dst: np.empty(n_rows, dtype=float) for dst in dst_map.values()}

            for pos in range(n_rows):
                ix = g_ix[pos]
                for dst in dst_map.values():
                    out_blocks[dst][pos] = float(np.median(collected[dst])) if collected[dst] else np.nan

                for src, dst in dst_map.items():
                    v = pd.to_numeric(df.loc[ix, src], errors='coerce')
                    if pd.notna(v):
                        collected[dst].append(float(v))

            for dst in dst_map.values():
                ser_ord = pd.Series(out_blocks[dst], index=g_ix)
                df.loc[g.index, dst] = ser_ord.reindex(g.index)

    @staticmethod
    def _variable_key_dataframe_for_curve_attachment(enriched_df):
        from App.SDM.Documenting.report_guide import report_guide

        cf_desc = report_guide.curve_feature_descriptions
        day_desc = report_guide.day_feature_descriptions
        rows = []
        for col in enriched_df.columns:
            col_s = str(col)
            if col_s == 'annotation_drinking_day':
                rows.append((col_s, day_desc.get(col_s, '')))
                continue
            m = re.match(r'^curve_(\d+)_(.+)$', col_s)
            if m:
                slot, base = m.group(1), m.group(2)
                txt = cf_desc.get(base)
                if txt is None:
                    stripped = base.replace('_CURVE', '')
                    txt = cf_desc.get(stripped)
                txt = txt or ''
                rows.append(
                    (
                        col_s,
                        f'Overlapping curve slot {slot} (chronological; join on SubID, Dataset_ID, curve_{slot}_id): {txt}',
                    )
                )
            else:
                rows.append((col_s, day_desc.get(col_s, '')))
        return pd.DataFrame(rows, columns=['Variable', 'Description'])

    def build_day_features_with_curve_attachments(
        self,
        curve_features_df,
        curve_feature_columns=None,
    ):
        """
        Copy ``self.day_features``, add ``annotation_drinking_day`` (blank strings for manual/review use),
        and ensure ``median_prior_*`` columns are present (recomputes if absent).

        Per-slot curve features (``curve_{n}_*``) and ``median_prior_*`` are expected to
        already exist on ``self.day_features`` from ``attach_curves_and_predict_drinking_days``.  This
        method is a thin wrapper that adds the annotation column and reorders columns for
        the workbook export.

        Args:
            curve_features_df: DataFrame from ``curveFeatures.curve_features`` (or compatible).
                Kept in signature for backwards compatibility; no longer used for merging.
            curve_feature_columns: Unused (kept for API compatibility).

        Returns:
            Enriched DataFrame (does not modify ``self.day_features``).
        """
        if curve_features_df is None or getattr(curve_features_df, 'empty', True):
            raise ValueError('curve_features_df must be a non-empty DataFrame')

        out = self.day_features.copy()
        original_cols_order = list(out.columns)
        out['annotation_drinking_day'] = ''
        # Ensure median_prior columns exist (no-op if already computed by attach_curves_and_predict_drinking_days)
        if not any(c.startswith('median_prior_') for c in out.columns):
            dayFeatures.assign_median_prior_from_all_qualifying_curves(out)
            dayFeatures.assign_median_prior_day_region_tac(out)
        return self._reorder_curve_attachment_columns(out, original_cols_order)

    def _reorder_curve_attachment_columns(self, df, original_cols_order):
        """Place ``annotation_drinking_day`` and ``median_prior_*`` near overlap blocks."""
        orig_set = set(original_cols_order)
        ann = 'annotation_drinking_day'
        attach_cols = [c for c in df.columns if c not in orig_set and c != ann]

        attach_by_slot: dict[int, list[str]] = {}
        median_prior_cols: list[str] = []
        other_attach: list[str] = []
        for c in attach_cols:
            m = re.match(r'^curve_(\d+)_(.+)$', str(c))
            if m:
                attach_by_slot.setdefault(int(m.group(1)), []).append(c)
            elif str(c).startswith('median_prior_'):
                median_prior_cols.append(c)
            else:
                other_attach.append(c)
        for k in attach_by_slot:
            attach_by_slot[k].sort()
        median_prior_cols.sort()

        ordered = list(original_cols_order)
        if ann in df.columns:
            anchor_ann = 'total_curve_overlap_hours'
            if anchor_ann in ordered:
                ix = ordered.index(anchor_ann) + 1
                ordered.insert(ix, ann)
            else:
                ordered.append(ann)
            ix_ann = ordered.index(ann)
            insert_at = ix_ann + 1
            for c in median_prior_cols:
                ordered.insert(insert_at, c)
                insert_at += 1

        for n in sorted(attach_by_slot.keys()):
            tail = attach_by_slot[n]
            anchor = f'curve_{n}_extends_next_day'
            if anchor in ordered:
                ix = ordered.index(anchor) + 1
                for c in reversed(tail):
                    ordered.insert(ix, c)
            else:
                ordered.extend([c for c in tail if c not in ordered])

        for c in other_attach:
            if c not in ordered:
                ordered.append(c)
        for c in df.columns:
            if c not in ordered:
                ordered.append(c)
        return df[[c for c in ordered if c in df.columns]]

    def export_day_level_with_curve_features_workbook(
        self,
        output_path,
        curve_features_df,
        curve_feature_columns=None,
        plot_column='signal_processing_plot',
        burst_ids=(1, 2, 3),
        only_inside_burst=True,
        subid_min=None,
        subid_max=None,
        row_interval=20,
        column_interval=12,
        x_scale=65 / 140,
        y_scale=90 / 182,
    ):
        """
        Build an enriched day-level copy (see ``build_day_features_with_curve_attachments``) and export:

        - **Features**: wide day table including ``annotation_drinking_day`` (empty until filled),
          per-slot curve features (``curve_{n}_*`` for slots 1–5, chronological),
          ``median_prior_*`` columns (medians of qualifying prior curves within burst),
          and any morning / self-report columns already present on ``self.day_features`` (e.g. after ``add_morning_report_drink_agreement``).
        - **Variable Key**: day descriptions from ``report_guide`` plus prefixed curve-slot descriptions.
        - **Day_plots_by_burst**: single worksheet stacking burst sections vertically — Burst 1 days 1–14 on top,
          then Burst 2, then Burst 3 (each block: days down rows, SubIDs across columns ascending).
          Row above each plot: ``annotation_drinking_day`` when filled. Morning self-report
          (``self-report: drinking | no drinking | missing``) is in the plot cell under the image.

        Args:
            output_path: Output ``.xlsx`` path.
            curve_features_df: Curve features DataFrame (same as used for overlap detection).
            curve_feature_columns: Optional explicit curve columns to attach; default uses stats + extras.
            plot_column: Day-level image path column for burst grids.
            burst_ids: Which burst IDs appear as stacked sections (top to bottom).
            only_inside_burst: If True and ``inside_burst`` exists, grids use rows with ``inside_burst == 1``.
            subid_min / subid_max: If set, export only rows with ``subid_min`` <= SubID <= ``subid_max``.
            row_interval / column_interval / x_scale / y_scale: Layout (matches other day exports).

        Returns:
            The enriched DataFrame written to the Features sheet.
        """
        enriched = self.build_day_features_with_curve_attachments(
            curve_features_df,
            curve_feature_columns=curve_feature_columns,
        )
        enriched = self.filter_day_features_by_subid_range(enriched, subid_min, subid_max)

        out_dir = os.path.dirname(output_path)
        if out_dir:
            os.makedirs(out_dir, exist_ok=True)

        drop_from_features = ['quality_analyzer', 'device_removal_plot']
        feat_sheet = enriched.drop(columns=[c for c in drop_from_features if c in enriched.columns], errors='ignore')
        vk = self._variable_key_dataframe_for_curve_attachment(feat_sheet)

        df_grid = enriched.copy()
        if only_inside_burst and 'inside_burst' in df_grid.columns:
            df_grid = df_grid[pd.to_numeric(df_grid['inside_burst'], errors='coerce') == 1].copy()
            print(f"  Burst plot grids: using inside_burst==1 only ({len(df_grid)} rows)")
        elif only_inside_burst:
            print("  Warning: inside_burst not found; burst grids include all day rows")

        print(f"\nExporting day-level + curve-feature workbook: {output_path}")

        with pd.ExcelWriter(output_path, engine='xlsxwriter', mode='w') as writer:
            book = writer.book
            bold = book.add_format({'bold': True})
            note_fmt = book.add_format({'text_wrap': True, 'valign': 'top'})
            lbl_fmt = book.add_format({'text_wrap': True, 'valign': 'bottom', 'font_size': 9})
            sr_under_plot_fmt = book.add_format({
                'text_wrap': True,
                'valign': 'bottom',
                'font_size': 8,
                'font_color': '#808080',
            })

            feat_sheet.to_excel(writer, sheet_name='Features', index=False)
            vk.to_excel(writer, sheet_name='Variable Key', index=False)
            ws_vk = writer.sheets['Variable Key']
            ws_vk.set_column(0, 0, 36, note_fmt)
            ws_vk.set_column(1, 1, 100, note_fmt)

            ann_col = 'annotation_drinking_day'
            sheet_name = 'Day_plots_by_burst'
            ws = book.add_worksheet(sheet_name)
            ws.set_column(0, 0, 56)
            ws.set_column(1, 1, 3)

            # Vertical stack: Burst 1 block, gap, Burst 2, gap, Burst 3, ...
            section_top = 0
            gap_between_bursts = 3

            for burst_id in burst_ids:
                bid = int(burst_id)
                lb = df_grid[pd.to_numeric(df_grid['Dataset_ID'], errors='coerce') == bid].copy()

                if lb.empty:
                    ws.write(section_top, 0, f'No day-level rows for burst {burst_id}', bold)
                    section_top += 2 + gap_between_bursts
                    continue

                subids = sorted(
                    pd.to_numeric(lb['SubID'], errors='coerce').dropna().astype(int).unique().tolist()
                )

                ws.write(section_top + 0, 0, f'Burst {burst_id}: day-level plots (SubID ascending)', bold)
                ws.write(
                    section_top + 1,
                    0,
                    'Ordering: SubIDs increase left to right. Row above plot: annotation_drinking_day. '
                    'Self-report label is in the plot cell under the image.',
                    note_fmt,
                )

                header_row = section_top + 2
                first_plot_row = section_top + 4
                for j, sid in enumerate(subids):
                    col_start = 2 + j * column_interval
                    ws.write(header_row, col_start, f'SubID {sid}', bold)

                for d in range(1, 15):
                    r = first_plot_row + (d - 1) * row_interval
                    ws.write(r, 0, f'Day {d}', bold)
                    for j, sid in enumerate(subids):
                        col_start = 2 + j * column_interval
                        rows_d = lb[
                            (pd.to_numeric(lb['SubID'], errors='coerce') == sid)
                            & (pd.to_numeric(lb['day_no'], errors='coerce') == d)
                        ]
                        if ann_col in rows_d.columns and not rows_d.empty:
                            v = rows_d[ann_col].iloc[0]
                            cell_txt = '' if pd.isna(v) else str(v)
                            ws.write(r - 1, col_start, cell_txt, lbl_fmt)
                        else:
                            ws.write(r - 1, col_start, '', lbl_fmt)

                        if rows_d.empty:
                            ws.write(r, col_start, '—')
                            continue

                        sr_lbl = self.self_report_plot_cell_label(rows_d.iloc[0])
                        ws.write(r, col_start, sr_lbl, sr_under_plot_fmt)

                        if plot_column not in rows_d.columns:
                            continue
                        raw_p = rows_d[plot_column].iloc[0]
                        if raw_p is None or (isinstance(raw_p, float) and np.isnan(raw_p)):
                            continue
                        p = str(raw_p).strip()
                        if not p or p.lower() == 'nan':
                            continue

                        if os.path.isfile(p):
                            try:
                                ws.insert_image(
                                    r,
                                    col_start,
                                    p,
                                    {'x_scale': x_scale, 'y_scale': y_scale},
                                )
                            except Exception:
                                pass

                # Next burst starts after day-14 block (14 row-interval steps) plus spacer
                section_top = first_plot_row + 14 * row_interval + gap_between_bursts

        print('  Done.')
        return enriched

    def export_workbook_days(self, file_name, split_plots_by=None, include_nonwear_plots=True, include_signal_processing_plots=True):
        """
        Export day features to an Excel workbook with plots.
        
        Args:
            file_name (str): Path to output Excel file
            split_plots_by (str, optional): How to split visualization tabs:
                - None (default): All plots in one 'Day Plots' tab
                - 'drinking_xgb': Split by pred_drinking_day (08.05 XGBoost)
            include_nonwear_plots (bool, optional): Whether to include non-wear detection plots (device_removal_plot). Default: True
            include_signal_processing_plots (bool, optional): Whether to include signal processing plots. Default: True
        """
        print("\nExporting workbook...")
        print(f"File name: {file_name}")
        print(f"Day features columns: {self.day_features.columns.tolist()}")
        
        # Import report guide for variable key
        from App.SDM.Documenting.report_guide import report_guide
        
        with pd.ExcelWriter(file_name, engine='xlsxwriter', mode='w') as writer:
            # Add variable key
            variable_key = pd.DataFrame({
                'Variable': list(report_guide.day_feature_descriptions.keys()),
                'Description': list(report_guide.day_feature_descriptions.values())
            })
            variable_key.to_excel(writer, sheet_name='Variable Key', index=False)
            
            # Filter out unwanted columns from features export
            columns_to_exclude = ['quality_analyzer', 'device_removal_plot', 'signal_processing_plot']
            filtered_features = self.day_features.drop(columns=columns_to_exclude, errors='ignore')
            
            # Write the features sheet
            filtered_features.to_excel(writer, sheet_name='Features', index=False)
            
            # Write the stats sheets
            row_index = 0
            for i, frame in enumerate(self.day_stat_frames):
                frame.to_excel(writer, sheet_name='Stats', startrow=row_index)
                row_index += len(frame) + 2
            
            # Embed the plots (only if plot columns exist and are requested)
            # Determine which plots to include
            plot_columns_to_include = []
            if include_nonwear_plots and 'device_removal_plot' in self.day_features.columns:
                plot_columns_to_include.append('device_removal_plot')
            if include_signal_processing_plots and 'signal_processing_plot' in self.day_features.columns:
                plot_columns_to_include.append('signal_processing_plot')
            
            # Only embed plots if at least one type is requested
            if plot_columns_to_include:
                for col in plot_columns_to_include:
                    print(self.day_features[col])
                
                if split_plots_by in ('drinking', 'drinking_by_start'):
                    warnings.warn(
                        f"split_plots_by={split_plots_by!r} is no longer available. "
                        "Use split_plots_by='drinking_xgb' (pred_drinking_day).",
                        DeprecationWarning,
                        stacklevel=2,
                    )
                    split_plots_by = None

                if split_plots_by == 'drinking_xgb' and 'pred_drinking_day' in self.day_features.columns:
                    drinking_days = self.day_features[self.day_features['pred_drinking_day'] == 1]
                    non_drinking_days = self.day_features[self.day_features['pred_drinking_day'] == 0]

                    if not non_drinking_days.empty:
                        embed_graphs_into_workbook_tab(
                            writer.book,
                            [non_drinking_days[col].tolist() for col in plot_columns_to_include],
                            worksheet_name = 'Non-Drinking Days (XGB)',
                            plot_header_text = '',
                            missing_plot_path_text = 'No Plot Available'
                        )

                    if not drinking_days.empty:
                        embed_graphs_into_workbook_tab(
                            writer.book,
                            [drinking_days[col].tolist() for col in plot_columns_to_include],
                            worksheet_name = 'Drinking Days (XGB)',
                            plot_header_text = '',
                            missing_plot_path_text = 'No Plot Available'
                        )

                else:
                    # Default: all plots in one tab
                    embed_graphs_into_workbook_tab(
                        writer.book,
                        [self.day_features[col].tolist() for col in plot_columns_to_include],
                        worksheet_name = 'Day Plots',
                        plot_header_text = '',
                        missing_plot_path_text = 'No Plot Available'
                    )
            