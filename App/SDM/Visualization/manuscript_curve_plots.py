#!/usr/bin/env python3
"""
Manuscript-style TAC curve zoom plots (cohort-agnostic).

Loads PROCESSED SDP files for requested (subid, curve_id) pairs, subsets to
each curve window plus 2 h before/after (Curve.region), and writes narrow
PNGs of final processed TAC only.

Titles come from a flexible label config (dict / table / CLI), e.g. memory
day type → "Fragmentary Day". No date, curve ID, or generic "TAC Curve" title.

Usage (from repo root):
  python App/SDM/Visualization/manuscript_curve_plots.py \\
    --cohort ARC --processed-subdir Burst1 \\
    --from-selected Results/ARC_all_devices/ARC_blackout_exemplar_curves_selected.xlsx \\
    --label-preset memlev_day

  python App/SDM/Visualization/manuscript_curve_plots.py \\
    --cohort ARC --processed-subdir Burst1 \\
    --pairs 1106:2 1143:25 \\
    --label-table Results/ARC_all_devices/ARC_blackout_exemplar_curves_selected.xlsx \\
    --label-sheet Curves --label-col memlev --label-preset memlev_day

  # Explicit per-pair titles:
  python App/SDM/Visualization/manuscript_curve_plots.py \\
    --cohort ARC --processed-subdir Burst1 \\
    --label-pair 1106:2:Fragmentary Day 1143:25:En Bloc Day

  # Signal-processing style (no SubID/Date/Curve subtitle):
  python App/SDM/Visualization/manuscript_curve_plots.py \\
    --cohort ARC --processed-subdir Burst1 \\
    --style signal_processing --pairs 1106:2 \\
    --output-dir Results/ARC/manuscript_curve_plots/signal_processing
"""

from __future__ import annotations

import argparse
import json
import math
import sys
from pathlib import Path
from typing import Any, Dict, List, Optional, Sequence, Tuple, Union

import matplotlib

matplotlib.use('Agg')
import matplotlib.dates as mdates
import matplotlib.pyplot as plt
from matplotlib.offsetbox import AnchoredText
import pandas as pd
from PIL import Image

PROJECT_ROOT = Path(__file__).resolve().parents[3]
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

from App.SDM.Configuration.file_management import load
from App.SDM.Visualization.tac import plot_signal_processing

DEFAULT_DATASET_ID = '001'
DEFAULT_FIGSIZE = (6.0, 4.0)
DEFAULT_SIGNAL_FIGSIZE = (16.0, 7.0)
DEFAULT_DPI = 300
STYLE_MANUSCRIPT = 'manuscript'
STYLE_SIGNAL_PROCESSING = 'signal_processing'
PLOT_STYLES = (STYLE_MANUSCRIPT, STYLE_SIGNAL_PROCESSING)
DEFAULT_SIGNAL_TITLE = 'Processed TAC Data'
# Larger typography for manuscript signal-processing exports only
SIGNAL_PROCESSING_FONTS = {
    'axis_label_fontsize': 20,
    'axis_label_fontweight': 'bold',
    'tick_label_fontsize': 18,       # y-axis ticks
    'x_tick_label_fontsize': 14,     # slightly smaller than y ticks
    'annotation_fontsize': 20,       # plot_event_lines shrinks slightly when ≥3 labels
    'title_fontsize': 24,
    'title_stats_fontsize': 16,
    'legend_fontsize': 13,
    'time_tick_format': '%-I%p',     # e.g. 2PM
    'x_tick_rotation': 0,
    'legend_bbox_to_anchor': (1.14, 1.22),  # tuck further top-right
    'title_ha': 'left',
    'title_x': 0.0,
}

# Compact FLAG_* → subtitle labels (aligned with manuscript Table 2 wording)
FLAG_SUBTITLE_LABELS = {
    'FLAG_unimputed_low_quality_curve': 'Non-imputed LQ (curve)',
    'FLAG_imputed_limit_curve': 'Imputed (curve)',
    'FLAG_unimputed_jump_curve': 'Unimputed jump (curve)',
    'FLAG_below_threshold_curve': 'Below threshold (curve)',
    'FLAG_incomplete_curve_start_curve': 'Incomplete rise',
    'FLAG_imputed_rise_curve': 'Imputed rise',
    'FLAG_incomplete_curve_end_curve': 'Incomplete fall',
    'FLAG_imputed_fall_curve': 'Imputed fall',
    'FLAG_gaps_and_non_wear_periphery_before': 'Gaps/non-wear (before)',
    'FLAG_extreme_negative_periphery_before': 'Very negative (before)',
    'FLAG_unimputed_low_quality_periphery_before': 'Non-imputed LQ (before)',
    'FLAG_gaps_and_non_wear_periphery_after': 'Gaps/non-wear (after)',
    'FLAG_extreme_negative_periphery_after': 'Very negative (after)',
    'FLAG_unimputed_low_quality_periphery_after': 'Non-imputed LQ (after)',
}

# Stratum slug → figure title for LQ example exports
SIGNAL_PROCESSING_TITLES = {
    'missing': 'TAC Curve: Imputed Missing Data Example',
    'non_wear': 'TAC Curve: Imputed Non-Wear Example',
    'jumps': 'TAC Curve: Imputed Jump Example',
    'plummets': 'TAC Curve: Imputed Plummet Example',
    'extreme_negative': 'TAC Curve: Imputed Very Negative Example',
}

# Shorter titles (no "TAC Curve:" prefix) for no-legend manuscript variants
SIGNAL_PROCESSING_TITLES_SHORT = {
    'missing': 'Imputed Missing Data Example',
    'non_wear': 'Imputed Non-Wear Example',
    'jumps': 'Imputed Jump Example',
    'plummets': 'Imputed Plummet Example',
    'extreme_negative': 'Imputed Very Negative Example',
}


def signal_processing_title(stratum: Optional[str] = None, short: bool = False) -> str:
    table = SIGNAL_PROCESSING_TITLES_SHORT if short else SIGNAL_PROCESSING_TITLES
    if stratum and stratum in table:
        return table[stratum]
    return DEFAULT_SIGNAL_TITLE


def drinks_from_annotations(annotations: Optional[Dict[str, Any]]) -> Optional[float]:
    """Parse drink count from raw keys like drinkFinish_*_6.0drks."""
    if not annotations:
        return None
    for key in annotations:
        text = str(key)
        if 'drks' not in text:
            continue
        if 'drinkStart' not in text and 'drinkFinish' not in text:
            continue
        try:
            return float(text.split('_')[-1].replace('drks', ''))
        except (TypeError, ValueError):
            continue
    return None


def flag_column_short_label(flag_col: str) -> str:
    """Map a FLAG_* column name to a short subtitle label."""
    if flag_col in FLAG_SUBTITLE_LABELS:
        return FLAG_SUBTITLE_LABELS[flag_col]
    raw = flag_col[5:] if flag_col.startswith('FLAG_') else flag_col
    return raw.replace('_', ' ').strip()


def active_flag_labels(
    row: Any,
    flag_cols: Optional[Sequence[str]] = None,
    *,
    label_map: Optional[Dict[str, str]] = None,
) -> List[str]:
    """Return short labels for FLAG_* columns that are active (truthy) on ``row``."""
    if row is None:
        return []
    labels_lookup = label_map if label_map is not None else FLAG_SUBTITLE_LABELS
    if flag_cols is None:
        try:
            cols = [c for c in row.index if str(c).startswith('FLAG_')]
        except AttributeError:
            cols = [c for c in getattr(row, 'keys', lambda: [])() if str(c).startswith('FLAG_')]
    else:
        cols = list(flag_cols)
    out: List[str] = []
    for col in cols:
        try:
            val = row[col]
        except Exception:
            continue
        if val is None or (isinstance(val, float) and pd.isna(val)) or pd.isna(val):
            continue
        try:
            active = bool(int(float(val)))
        except (TypeError, ValueError):
            active = bool(val)
        if not active:
            continue
        out.append(labels_lookup.get(col, flag_column_short_label(col)))
    return out


def format_flagging_subtitle(
    row: Any,
    flag_cols: Optional[Sequence[str]] = None,
    *,
    prefix: str = 'Flags: ',
    sep: str = '; ',
    empty: str = '',
) -> str:
    """Build an italic subtitle listing active flagging reasons for a pairs row."""
    labels = active_flag_labels(row, flag_cols)
    if not labels:
        return empty
    return f'{prefix}{sep.join(labels)}'


def format_title_stats(
    drinks: Optional[Any] = None,
    ebac: Optional[Any] = None,
) -> str:
    """Build 'Drinks: …  |  eBAC: …' line for manuscript signal-processing titles."""
    parts: List[str] = []
    if drinks is not None and pd.notna(drinks):
        try:
            d = float(drinks)
            parts.append(f'Drinks: {d:g}')
        except (TypeError, ValueError):
            parts.append(f'Drinks: {drinks}')
    if ebac is not None and pd.notna(ebac):
        try:
            parts.append(f'eBAC: {float(ebac):.3f}')
        except (TypeError, ValueError):
            parts.append(f'eBAC: {ebac}')
    return '  |  '.join(parts)

Pair = Tuple[int, int]
LabelLookup = Dict[Pair, str]

# Preset: raw memlev / memlev_label → manuscript day-type title
MEMLEV_DAYTYPE_LABELS: Dict[Any, str] = {
    0: 'No Memory Loss',
    0.0: 'No Memory Loss',
    1: 'Memory Loss: Fragmentary',
    1.0: 'Memory Loss: Fragmentary',
    2: 'Memory Loss: En Bloc',
    2.0: 'Memory Loss: En Bloc',
    '0': 'No Memory Loss',
    '1': 'Memory Loss: Fragmentary',
    '2': 'Memory Loss: En Bloc',
    'No memory loss': 'No Memory Loss',
    'Fragmentary': 'Memory Loss: Fragmentary',
    'En bloc': 'Memory Loss: En Bloc',
    'En Bloc': 'Memory Loss: En Bloc',
}

LABEL_PRESETS = {
    'memlev_day': MEMLEV_DAYTYPE_LABELS,
}

# Composite grid: rows top→bottom, columns = persons
COMPOSITE_ROW_ORDER = [
    'Memory Loss: En Bloc',
    'Memory Loss: Fragmentary',
    'No Memory Loss',
]


def resolve_processed_dir(
    cohort: str,
    processed_subdir: Optional[str] = None,
    project_root: Path = PROJECT_ROOT,
) -> Path:
    """Return Inputs/Skyn_Data_PROCESSED/{cohort}/[{subdir}]."""
    base = project_root / 'Inputs' / 'Skyn_Data_PROCESSED' / cohort
    path = base / processed_subdir if processed_subdir else base
    if not path.is_dir():
        raise FileNotFoundError(f'PROCESSED directory not found: {path}')
    return path


def resolve_output_dir(
    cohort: str,
    output_dir: Optional[Path] = None,
    project_root: Path = PROJECT_ROOT,
    style: str = STYLE_MANUSCRIPT,
) -> Path:
    if output_dir is not None:
        out = output_dir if output_dir.is_absolute() else project_root / output_dir
    elif style == STYLE_SIGNAL_PROCESSING:
        out = (
            project_root / 'Results' / cohort / 'manuscript_curve_plots'
            / 'signal_processing'
        )
    else:
        out = project_root / 'Results' / cohort / 'manuscript_curve_plots'
    out.mkdir(parents=True, exist_ok=True)
    return out


def load_skyn_dataset(
    processed_dir: Path,
    subid: int,
    dataset_id: str = DEFAULT_DATASET_ID,
):
    """Load `{subid}_{dataset_id}_skyn_data_processed.sdp` from processed_dir."""
    try:
        dataset_id = f'{int(float(dataset_id)):03d}'
    except (TypeError, ValueError):
        dataset_id = str(dataset_id)
    name = f'{int(subid)}_{dataset_id}_skyn_data_processed.sdp'
    sdm_path = processed_dir / f'{name}.sdm'
    pickle_path = processed_dir / f'{name}.pickle'
    if not sdm_path.is_file() and not pickle_path.is_file():
        raise FileNotFoundError(
            f'Processed SDP not found for subid={subid} dataset_id={dataset_id} '
            f'under {processed_dir} (tried {name}.sdm / .pickle)'
        )
    return load(name, str(processed_dir))


def find_curve(skyn_dataset, curve_id: int):
    """Return Curve object matching curve_id, or None."""
    curves = getattr(skyn_dataset, 'curves', None) or []
    for curve in curves:
        if int(curve.curve_id) == int(curve_id):
            return curve
    return None


def parse_pairs(pair_strings: Sequence[str]) -> List[Pair]:
    """Parse 'subid:curve_id' tokens into (subid, curve_id) tuples."""
    pairs: List[Pair] = []
    for token in pair_strings:
        text = str(token).strip()
        if not text:
            continue
        if ':' not in text:
            raise ValueError(f'Invalid pair {token!r}; expected subid:curve_id')
        left, right = text.split(':', 1)
        pairs.append((int(left), int(right)))
    return pairs


def parse_label_pairs(tokens: Sequence[str]) -> LabelLookup:
    """Parse 'subid:curve_id:Title With Spaces' into a label lookup."""
    lookup: LabelLookup = {}
    for token in tokens:
        text = str(token).strip()
        parts = text.split(':', 2)
        if len(parts) != 3:
            raise ValueError(
                f'Invalid --label-pair {token!r}; expected subid:curve_id:Title'
            )
        lookup[(int(parts[0]), int(parts[1]))] = parts[2].strip()
    return lookup


def _read_pair_table(path: Path, sheet: Optional[str] = None) -> pd.DataFrame:
    path = Path(path)
    if not path.is_file():
        raise FileNotFoundError(f'Table not found: {path}')
    if path.suffix.lower() == '.csv':
        return pd.read_csv(path)

    xl = pd.ExcelFile(path)
    if sheet and sheet in xl.sheet_names:
        return pd.read_excel(path, sheet_name=sheet)
    # Prefer Selected, else Curves filtered to selected==1, else first sheet
    if 'Selected' in xl.sheet_names:
        return pd.read_excel(path, sheet_name='Selected')
    if 'Curves' in xl.sheet_names:
        df = pd.read_excel(path, sheet_name='Curves')
        if 'selected' in df.columns:
            return df[df['selected'] == 1].copy()
        return df
    return pd.read_excel(path, sheet_name=xl.sheet_names[0])


def pairs_from_selected_workbook(path: Path) -> List[Pair]:
    """Read Selected tab (or Curves with selected==1) for (subid, curve_id)."""
    df = _read_pair_table(path)
    if 'subid' not in df.columns or 'curve_id' not in df.columns:
        raise ValueError(f'{path} missing subid/curve_id columns')
    return [
        (int(r.subid), int(r.curve_id))
        for r in df[['subid', 'curve_id']].drop_duplicates().itertuples(index=False)
    ]


def resolve_value_map(map_spec: Optional[Union[Dict[Any, str], str]]) -> Dict[Any, str]:
    """Accept a dict, preset name, or JSON object/string."""
    if map_spec is None:
        return {}
    if isinstance(map_spec, dict):
        return dict(map_spec)
    text = str(map_spec).strip()
    if text in LABEL_PRESETS:
        return dict(LABEL_PRESETS[text])
    # JSON file or inline JSON
    path = Path(text)
    if path.is_file():
        with open(path, 'r') as f:
            loaded = json.load(f)
        if not isinstance(loaded, dict):
            raise ValueError(f'Label map JSON must be an object: {path}')
        return loaded
    loaded = json.loads(text)
    if not isinstance(loaded, dict):
        raise ValueError('Inline --label-map JSON must be an object')
    return loaded


def build_label_lookup(
    label_config: Optional[Dict[str, Any]] = None,
) -> LabelLookup:
    """
    Build (subid, curve_id) → title from a label config dict.

    Supported configs
    -----------------
    Direct lookup
      {'source': 'lookup', 'labels': {(1106, 2): 'Fragmentary Day', ...}}
      or {'source': 'lookup', 'labels': {'1106:2': 'Fragmentary Day', ...}}

    Table (Excel/CSV)
      {
        'source': 'table',
        'path': 'Results/.../file.xlsx',
        'sheet': 'Selected',          # optional
        'key_cols': ['subid', 'curve_id'],
        'value_col': 'memlev',        # raw values mapped via 'map'
        'map': 'memlev_day' | {...},  # optional remapping
        'filter_col': 'selected',     # optional
        'filter_value': 1,
        'default': '',
      }

    None / empty → {}
    """
    if not label_config:
        return {}

    source = str(label_config.get('source', 'lookup')).lower()
    default = label_config.get('default', '')

    if source == 'lookup':
        raw = label_config.get('labels') or label_config.get('map') or {}
        lookup: LabelLookup = {}
        for key, title in raw.items():
            if isinstance(key, tuple) and len(key) == 2:
                lookup[(int(key[0]), int(key[1]))] = str(title)
            elif isinstance(key, str) and ':' in key:
                a, b = key.split(':', 1)
                lookup[(int(a), int(b))] = str(title)
            else:
                raise ValueError(
                    f'lookup label key must be (subid, curve_id) or "subid:curve_id"; got {key!r}'
                )
        return lookup

    if source == 'table':
        path = Path(label_config['path'])
        if not path.is_absolute():
            path = PROJECT_ROOT / path
        sheet = label_config.get('sheet')
        df = _read_pair_table(path, sheet=sheet)
        filter_col = label_config.get('filter_col')
        if filter_col and filter_col in df.columns:
            df = df[df[filter_col] == label_config.get('filter_value', 1)].copy()

        key_cols = list(label_config.get('key_cols', ['subid', 'curve_id']))
        value_col = label_config.get('value_col', 'memlev')
        for col in key_cols + [value_col]:
            if col not in df.columns:
                raise ValueError(f'Label table missing column {col!r} in {path}')

        value_map = resolve_value_map(label_config.get('map'))
        lookup = {}
        for _, row in df.iterrows():
            pair = (int(row[key_cols[0]]), int(row[key_cols[1]]))
            raw_val = row[value_col]
            if value_map:
                title = value_map.get(raw_val, value_map.get(str(raw_val), default))
                if title == default and pd.notna(raw_val) and raw_val not in value_map:
                    # try float/int coercion for memlev-like cols
                    try:
                        title = value_map.get(float(raw_val), default)
                    except (TypeError, ValueError):
                        pass
            else:
                title = default if pd.isna(raw_val) else str(raw_val)
            lookup[pair] = str(title) if title is not None else str(default)
        return lookup

    if source in ('none', ''):
        return {}

    raise ValueError(
        f"Unknown label_config source {source!r}; use 'lookup', 'table', or 'none'"
    )


def plot_manuscript_curve(
    df: pd.DataFrame,
    out_path: Path,
    *,
    tac_column: str = 'TAC',
    curve_threshold: Optional[float] = None,
    peak: Optional[float] = None,
    duration_hours: Optional[float] = None,
    title: str = '',
    figsize: Tuple[float, float] = DEFAULT_FIGSIZE,
    dpi: int = DEFAULT_DPI,
    show_peak: bool = True,
    ylim: Optional[Tuple[float, float]] = None,
    ytick_step: Optional[float] = None,
) -> Path:
    """
    Narrow manuscript PNG: final processed TAC only (line + markers).

    ``title`` is the sole figure title (e.g. "Fragmentary Day"). No subtitle.
    When ``ylim`` is set (standardized axis), all plots share the same range;
    ``ytick_step`` controls major y tick spacing (e.g. 100 → …, 200, 300).
    Peak and duration (hours) are annotated in the top-right when provided.
    """
    if df is None or df.empty:
        raise ValueError('Empty curve dataframe')
    if tac_column not in df.columns:
        raise ValueError(f'TAC column {tac_column!r} not in curve data')
    if 'datetime' not in df.columns:
        raise ValueError("Curve data missing 'datetime' column")

    work = df.copy()
    work['datetime'] = pd.to_datetime(work['datetime'])
    tac = pd.to_numeric(work[tac_column], errors='coerce')

    fig, ax = plt.subplots(figsize=figsize)
    ax.plot(work['datetime'], tac, color='black', linewidth=1.5, zorder=2)
    ax.scatter(
        work['datetime'], tac, color='black', s=8, zorder=3, linewidths=0,
    )

    t_min = work['datetime'].min()
    t_max = work['datetime'].max()
    if curve_threshold is not None and pd.notna(curve_threshold):
        ax.hlines(
            float(curve_threshold),
            xmin=t_min,
            xmax=t_max,
            colors='black',
            linestyles='--',
            linewidth=1.0,
            zorder=1,
        )

    if show_peak and peak is not None and pd.notna(peak):
        peak_rows = work.loc[tac == float(peak), 'datetime']
        if peak_rows.empty:
            idx = (tac - float(peak)).abs().idxmin()
            peak_time = work.loc[idx, 'datetime']
            peak_y = float(tac.loc[idx])
        else:
            peak_time = peak_rows.iloc[0]
            peak_y = float(peak)
        y0 = (
            float(curve_threshold)
            if curve_threshold is not None and pd.notna(curve_threshold)
            else float(tac.min())
        )
        ax.vlines(
            peak_time,
            ymin=y0,
            ymax=peak_y,
            colors='black',
            linestyles='--',
            linewidth=1.0,
            zorder=1,
        )

    ax.set_xlabel('Time', fontsize=11)
    ax.set_ylabel('TAC', fontsize=11)
    if title:
        ax.set_title(title, fontsize=13, fontweight='semibold', pad=10)

    ax.tick_params(axis='both', labelsize=9)
    ax.xaxis.set_major_locator(mdates.HourLocator(interval=2))
    ax.xaxis.set_major_formatter(mdates.DateFormatter('%-I%p'))  # e.g. 2PM
    fig.autofmt_xdate(rotation=30, ha='right')

    if ylim is not None:
        ax.set_ylim(float(ylim[0]), float(ylim[1]))
        step = ytick_step
        if step is None:
            span = float(ylim[1]) - float(ylim[0])
            if span > 200:
                step = 100.0
            elif span > 50:
                step = 50.0
            else:
                step = 10.0
        ax.yaxis.set_major_locator(plt.MultipleLocator(step))
    else:
        y_max = float(tac.max()) if tac.notna().any() else 0.0
        y_min = float(tac.min()) if tac.notna().any() else 0.0
        if y_max < 10:
            ax.set_ylim(-5, 20)
        else:
            pad = max(5.0, 0.05 * (y_max - y_min))
            ax.set_ylim(y_min - pad, y_max + pad)

    # Top-right stats box: peak + curve duration (left-justified, flush to corner)
    stats_lines = []
    if peak is not None and pd.notna(peak):
        stats_lines.append(f'Peak: {float(peak):.1f} μg/L')
    if duration_hours is not None and pd.notna(duration_hours):
        stats_lines.append(f'Duration: {float(duration_hours):.1f} h')
    if stats_lines:
        box = AnchoredText(
            '\n'.join(stats_lines),
            loc='upper right',
            prop={'size': 9},
            frameon=True,
            borderpad=0.35,
            pad=0.35,
        )
        box.patch.set_boxstyle('round,pad=0.35')
        box.patch.set_facecolor('white')
        box.patch.set_edgecolor('black')
        box.patch.set_linewidth(0.8)
        box.patch.set_alpha(0.92)
        ax.add_artist(box)

    out_path = Path(out_path)
    out_path.parent.mkdir(parents=True, exist_ok=True)
    plt.tight_layout()
    fig.savefig(out_path, dpi=dpi, bbox_inches='tight')
    plt.close(fig)
    return out_path


def _curve_peak_and_tac_extent(curve) -> Tuple[float, float, float]:
    """Return (peak, tac_min_in_region, tac_max_in_region) for a Curve object."""
    tac_column = getattr(curve, 'TAC_column', 'TAC')
    peak = None
    if hasattr(curve, 'curve_tac_features') and isinstance(curve.curve_tac_features, dict):
        peak = curve.curve_tac_features.get('peak_CURVE')
    region_tac = pd.to_numeric(curve.region[tac_column], errors='coerce')
    if peak is None:
        curve_tac = pd.to_numeric(curve.curve[tac_column], errors='coerce')
        peak = float(curve_tac.max()) if curve_tac.notna().any() else float('nan')
    else:
        peak = float(peak)
    tac_min = float(region_tac.min()) if region_tac.notna().any() else 0.0
    tac_max = float(region_tac.max()) if region_tac.notna().any() else peak
    return peak, tac_min, tac_max


def compute_standardized_ylim(
    peaks: Sequence[float],
    tac_mins: Optional[Sequence[float]] = None,
    *,
    ylim_max: Optional[float] = None,
    pad_frac: float = 0.05,
) -> Tuple[Tuple[float, float], float, float]:
    """
    Shared y-limits from the batch's largest peak.

    Returns ((y0, y1), batch_peak_max, ytick_step).
    ``ytick_step`` is chosen so the top major tick is a round value near the
    peak (e.g. peak≈301 → step 100 → top tick 300, axis slightly above peak).
    """
    peak_vals = [float(p) for p in peaks if p is not None and pd.notna(p)]
    if not peak_vals and ylim_max is None:
        return (0.0, 20.0), 0.0, 10.0

    batch_peak = max(peak_vals) if peak_vals else float(ylim_max)
    if ylim_max is not None:
        y1 = float(ylim_max)
    else:
        y1 = batch_peak * (1.0 + pad_frac)

    mins = [float(v) for v in (tac_mins or []) if v is not None and pd.notna(v)]
    y0 = min(0.0, min(mins)) if mins else 0.0

    span = y1 - y0
    if span > 200:
        step = 100.0
    elif span > 50:
        step = 50.0
    else:
        step = 10.0

    # Ensure axis top clears the peak and sits at/above a major tick
    top_tick = math.floor(batch_peak / step) * step
    if top_tick < batch_peak:
        # keep top tick at floor (e.g. 300 for peak 301); axis goes a bit higher
        y1 = max(y1, batch_peak * (1.0 + pad_frac), top_tick + step * 0.15)
    if ylim_max is not None:
        y1 = float(ylim_max)

    return (y0, y1), batch_peak, step


def export_pair(
    processed_dir: Path,
    output_dir: Path,
    subid: int,
    curve_id: int,
    dataset_id: str = DEFAULT_DATASET_ID,
    figsize: Tuple[float, float] = DEFAULT_FIGSIZE,
    show_peak: bool = True,
    title: str = '',
    ylim: Optional[Tuple[float, float]] = None,
    ytick_step: Optional[float] = None,
    dataset_cache: Optional[Dict[Tuple[int, str], object]] = None,
    style: str = STYLE_MANUSCRIPT,
    drinks: Optional[Any] = None,
    ebac: Optional[Any] = None,
    clean_event_labels: Optional[bool] = None,
    signal_plot_overrides: Optional[Dict[str, Any]] = None,
    subtitle_text: str = '',
    pin_title: Optional[bool] = None,
) -> Dict[str, object]:
    """Load SDP (cached), find curve, write manuscript or signal-processing PNG.

    For ``style='signal_processing'``, event labels are cleaned (Drink Start/End)
    and Drinks/eBAC are shown under the title unless ``clean_event_labels=False``.
    ``skyn_dataset`` continues to call ``plot_signal_processing`` with defaults
    (raw labels, no title_stats).
    signal_plot_overrides : optional kwargs forwarded to ``plot_signal_processing``
    (e.g. ``show_legend=False``, ``filename_suffix='_no_lgnd'``, centered titles).
    subtitle_text : optional middle line (e.g. flagging reasons) under the title.
    pin_title : if True, force pinned manuscript header; if None, auto when
    ``subtitle_text`` or title_stats is non-empty.
    """
    if style not in PLOT_STYLES:
        raise ValueError(f'Unknown style {style!r}; expected one of {PLOT_STYLES}')

    cache = dataset_cache if dataset_cache is not None else {}
    key = (int(subid), str(dataset_id))
    if key not in cache:
        print(f'Loading SDP subid={subid} dataset_id={dataset_id} ...')
        cache[key] = load_skyn_dataset(processed_dir, subid, dataset_id)
    skyn_dataset = cache[key]

    curve = find_curve(skyn_dataset, curve_id)
    if curve is None:
        available = [int(c.curve_id) for c in (skyn_dataset.curves or [])]
        raise KeyError(
            f'curve_id={curve_id} not found for subid={subid}; '
            f'available={available}'
        )

    # Curve demarcation ± 2 h periphery (Curve.region)
    df = curve.region
    tac_column = getattr(curve, 'TAC_column', 'TAC')
    peak = None
    duration_hours = None
    begin_curve = end_curve = None
    if hasattr(curve, 'curve_tac_features') and isinstance(curve.curve_tac_features, dict):
        peak = curve.curve_tac_features.get('peak_CURVE')
        duration_hours = curve.curve_tac_features.get('duration_CURVE')
        begin_curve = curve.curve_tac_features.get('begin_CURVE')
        end_curve = curve.curve_tac_features.get('end_CURVE')
    if peak is None and tac_column in curve.curve.columns:
        peak = float(pd.to_numeric(curve.curve[tac_column], errors='coerce').max())
    if duration_hours is None and begin_curve is not None and end_curve is not None:
        duration_hours = (
            (pd.Timestamp(end_curve) - pd.Timestamp(begin_curve)).total_seconds() + 60
        ) / 3600.0
    elif duration_hours is None:
        duration_hours = len(curve.curve) / 60.0

    window_begin = pd.Timestamp(df['datetime'].iloc[0])
    window_end = pd.Timestamp(df['datetime'].iloc[-1])

    output_dir = Path(output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)

    if style == STYLE_SIGNAL_PROCESSING:
        plot_title = title if title else DEFAULT_SIGNAL_TITLE
        annotations = getattr(curve, 'curve_plot_annotations', None) or {}
        use_clean = True if clean_event_labels is None else bool(clean_event_labels)
        drinks_val = drinks
        if drinks_val is None or (isinstance(drinks_val, float) and pd.isna(drinks_val)):
            drinks_val = drinks_from_annotations(annotations)
        stats = format_title_stats(drinks_val, ebac)
        # plot_signal_processing requires trailing slash on folder path
        plot_folder = str(output_dir)
        if not plot_folder.endswith(('/', '\\')):
            plot_folder = plot_folder + '/'
        plot_kwargs = dict(SIGNAL_PROCESSING_FONTS)
        if signal_plot_overrides:
            plot_kwargs.update(signal_plot_overrides)
        sub = (subtitle_text or '').strip()
        # Pin manuscript header when subtitle and/or title_stats need the 2–3 line layout
        if pin_title is None:
            use_pin = bool(sub) or bool(stats)
        else:
            use_pin = bool(pin_title)
        written = plot_signal_processing(
            df,
            plot_folder,
            int(subid),
            int(curve_id),
            dataset_id,
            'CURVE',
            getattr(curve, 'curve_threshold', None),
            event_timestamps=annotations,
            subtitle_text=sub,
            show_imputations=True,
            title=plot_title,
            clean_event_labels=use_clean,
            title_stats=stats,
            pin_title=use_pin,
            **plot_kwargs,
        )
        out_path = Path(written)
        # Ensure figsize is logged even though tac.py uses its own size
        used_figsize = DEFAULT_SIGNAL_FIGSIZE
        used_show_peak = False
    else:
        out_name = f'{int(subid)}_{dataset_id}_{int(curve_id)}_TAC_manuscript.png'
        out_path = output_dir / out_name
        plot_manuscript_curve(
            df,
            out_path,
            tac_column=tac_column,
            curve_threshold=getattr(curve, 'curve_threshold', None),
            peak=peak,
            duration_hours=duration_hours,
            title=title,
            figsize=figsize,
            show_peak=show_peak,
            ylim=ylim,
            ytick_step=ytick_step,
        )
        used_figsize = figsize
        used_show_peak = show_peak
        drinks_val = drinks
        stats = ''

    row = {
        'subid': int(subid),
        'dataset_id': dataset_id,
        'curve_id': int(curve_id),
        'title': title,
        'style': style,
        'begin_CURVE': begin_curve,
        'end_CURVE': end_curve,
        'window_begin': window_begin,
        'window_end': window_end,
        'peak': peak,
        'duration_hours': duration_hours,
        'n_points': int(len(df)),
        'tac_column': tac_column,
        'ylim': ylim if style == STYLE_MANUSCRIPT else None,
        'ytick_step': ytick_step if style == STYLE_MANUSCRIPT else None,
        'figsize': used_figsize,
        'show_peak': used_show_peak,
        'drinks': drinks_val,
        'ebac': ebac,
        'title_stats': stats if style == STYLE_SIGNAL_PROCESSING else '',
        'subtitle_text': (subtitle_text or '').strip() if style == STYLE_SIGNAL_PROCESSING else '',
        'path': str(out_path),
    }
    print(
        f'  Wrote {out_path.name}  style={style}  title={title!r}  '
        f'n={row["n_points"]}  window={window_begin} → {window_end}'
    )
    return row


def export_pairs(
    cohort: str,
    pairs: Sequence[Pair],
    processed_subdir: Optional[str] = None,
    dataset_id: str = DEFAULT_DATASET_ID,
    output_dir: Optional[Path] = None,
    figsize: Tuple[float, float] = DEFAULT_FIGSIZE,
    show_peak: bool = True,
    label_config: Optional[Dict[str, Any]] = None,
    label_lookup: Optional[LabelLookup] = None,
    standardize_y: bool = False,
    ylim_max: Optional[float] = None,
    project_root: Path = PROJECT_ROOT,
    style: str = STYLE_MANUSCRIPT,
) -> pd.DataFrame:
    if style not in PLOT_STYLES:
        raise ValueError(f'Unknown style {style!r}; expected one of {PLOT_STYLES}')

    processed_dir = resolve_processed_dir(cohort, processed_subdir, project_root)
    out_dir = resolve_output_dir(cohort, output_dir, project_root, style=style)
    print(f'PROCESSED dir: {processed_dir}')
    print(f'Output dir:    {out_dir}')
    print(f'Style:         {style}')
    print(f'Pairs: {list(pairs)}')

    lookup: LabelLookup = {}
    if label_lookup:
        lookup.update(label_lookup)
    if label_config:
        lookup.update(build_label_lookup(label_config))

    cache: Dict[Tuple[int, str], object] = {}
    ylim: Optional[Tuple[float, float]] = None
    ytick_step: Optional[float] = None

    if style == STYLE_MANUSCRIPT and (standardize_y or ylim_max is not None):
        peaks: List[float] = []
        tac_mins: List[float] = []
        for subid, curve_id in pairs:
            key = (int(subid), str(dataset_id))
            if key not in cache:
                print(f'Loading SDP subid={subid} dataset_id={dataset_id} ...')
                cache[key] = load_skyn_dataset(processed_dir, subid, dataset_id)
            curve = find_curve(cache[key], curve_id)
            if curve is None:
                continue
            peak, tac_min, _ = _curve_peak_and_tac_extent(curve)
            if pd.notna(peak):
                peaks.append(float(peak))
            tac_mins.append(float(tac_min))
        ylim, batch_peak, ytick_step = compute_standardized_ylim(
            peaks, tac_mins, ylim_max=ylim_max,
        )
        print(
            f'Standardized y-axis: ylim={ylim}  '
            f'batch_peak_max={batch_peak:.4g}  ytick_step={ytick_step}'
        )

    rows: List[Dict[str, object]] = []
    errors: List[str] = []
    for subid, curve_id in pairs:
        title = lookup.get((int(subid), int(curve_id)), '')
        try:
            rows.append(
                export_pair(
                    processed_dir,
                    out_dir,
                    subid,
                    curve_id,
                    dataset_id=dataset_id,
                    figsize=figsize,
                    show_peak=show_peak,
                    title=title,
                    ylim=ylim,
                    ytick_step=ytick_step,
                    dataset_cache=cache,
                    style=style,
                )
            )
        except Exception as exc:
            msg = f'subid={subid} curve_id={curve_id}: {exc}'
            print(f'  ERROR {msg}')
            errors.append(msg)

    manifest = pd.DataFrame(rows)
    if not manifest.empty:
        manifest_path = out_dir / 'manifest.csv'
        manifest.to_csv(manifest_path, index=False)
        print(f'Manifest: {manifest_path}')
    if errors:
        print(f'Completed with {len(errors)} error(s).')
    return manifest


def combine_manuscript_grid(
    manifest: pd.DataFrame,
    output_path: Path,
    *,
    row_titles: Optional[Sequence[str]] = None,
    col_subids: Optional[Sequence[int]] = None,
    pad_px: int = 12,
    bg_color: Tuple[int, int, int] = (255, 255, 255),
) -> Path:
    """
    Combine individual manuscript PNGs into one grid image.

    Columns = persons (subid), rows = memory-loss titles in order:
    En Bloc → Fragmentary → No Memory Loss (override via ``row_titles``).

    ``manifest`` must include columns ``subid``, ``title``, and ``path``.
    Missing cells are left blank (white).
    """
    if manifest is None or manifest.empty:
        raise ValueError('manifest is empty; nothing to combine')
    for col in ('subid', 'title', 'path'):
        if col not in manifest.columns:
            raise ValueError(f'manifest missing required column {col!r}')

    rows = list(row_titles) if row_titles is not None else list(COMPOSITE_ROW_ORDER)
    if col_subids is not None:
        cols = [int(s) for s in col_subids]
    else:
        # Preserve first-seen subid order from manifest
        cols = list(dict.fromkeys(int(s) for s in manifest['subid'].tolist()))

    # Resolve path grid
    path_grid: List[List[Optional[Path]]] = []
    for title in rows:
        row_paths: List[Optional[Path]] = []
        for subid in cols:
            hits = manifest[
                (manifest['subid'].astype(int) == int(subid))
                & (manifest['title'].astype(str) == str(title))
            ]
            if hits.empty:
                row_paths.append(None)
                print(f'  Grid warning: missing subid={subid} title={title!r}')
            else:
                row_paths.append(Path(str(hits.iloc[0]['path'])))
        path_grid.append(row_paths)

    # Load images; normalize cell size to max width/height among present plots
    images: List[List[Optional[Image.Image]]] = []
    max_w = max_h = 0
    for row_paths in path_grid:
        row_imgs: List[Optional[Image.Image]] = []
        for p in row_paths:
            if p is None or not p.is_file():
                row_imgs.append(None)
                continue
            im = Image.open(p).convert('RGB')
            row_imgs.append(im)
            max_w = max(max_w, im.width)
            max_h = max(max_h, im.height)
        images.append(row_imgs)

    if max_w == 0 or max_h == 0:
        raise ValueError('No plot images found on disk to combine')

    n_rows = len(rows)
    n_cols = len(cols)
    canvas_w = n_cols * max_w + (n_cols + 1) * pad_px
    canvas_h = n_rows * max_h + (n_rows + 1) * pad_px
    canvas = Image.new('RGB', (canvas_w, canvas_h), bg_color)

    for r, row_imgs in enumerate(images):
        for c, im in enumerate(row_imgs):
            x = pad_px + c * (max_w + pad_px)
            y = pad_px + r * (max_h + pad_px)
            if im is None:
                continue
            # Center within cell if sizes differ
            paste = im
            if im.width != max_w or im.height != max_h:
                cell = Image.new('RGB', (max_w, max_h), bg_color)
                ox = (max_w - im.width) // 2
                oy = (max_h - im.height) // 2
                cell.paste(im, (ox, oy))
                paste = cell
            canvas.paste(paste, (x, y))

    output_path = Path(output_path)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    canvas.save(output_path, format='PNG')
    print(
        f'Combined grid ({n_rows} rows × {n_cols} cols) → {output_path}'
    )
    return output_path


def parse_args(argv: Optional[Sequence[str]] = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description=(
            'Export narrow manuscript TAC PNGs for (subid, curve_id) pairs '
            'from cohort PROCESSED SDP files (±2 h window; flexible titles)'
        )
    )
    parser.add_argument(
        '--cohort',
        required=True,
        help='Cohort name under Inputs/Skyn_Data_PROCESSED/ and Results/',
    )
    parser.add_argument(
        '--processed-subdir',
        default=None,
        help='Optional nested folder under the cohort PROCESSED dir (e.g. Burst1)',
    )
    parser.add_argument(
        '--dataset-id',
        default=DEFAULT_DATASET_ID,
        help=f'Dataset identifier in SDP filename (default {DEFAULT_DATASET_ID})',
    )
    parser.add_argument(
        '--pairs',
        nargs='+',
        default=None,
        help='Pairs as subid:curve_id (e.g. 1106:2 1143:25)',
    )
    parser.add_argument(
        '--from-selected',
        type=Path,
        default=None,
        help='Excel workbook with Selected tab (or Curves with selected==1)',
    )
    parser.add_argument(
        '--output-dir',
        type=Path,
        default=None,
        help=(
            'Output folder (default Results/{cohort}/manuscript_curve_plots '
            'or .../manuscript_curve_plots/signal_processing when --style '
            'signal_processing)'
        ),
    )
    parser.add_argument(
        '--style',
        choices=list(PLOT_STYLES),
        default=STYLE_MANUSCRIPT,
        help=(
            'Plot style: manuscript (narrow B&W TAC) or signal_processing '
            '(full quality/imputation markers, no subtitle; default manuscript)'
        ),
    )
    parser.add_argument(
        '--no-peak',
        action='store_true',
        help='Omit peak vertical annotation',
    )
    parser.add_argument(
        '--figsize',
        nargs=2,
        type=float,
        default=list(DEFAULT_FIGSIZE),
        metavar=('W', 'H'),
        help=f'Figure size inches (default {DEFAULT_FIGSIZE[0]} {DEFAULT_FIGSIZE[1]})',
    )
    parser.add_argument(
        '--standardize-y',
        action='store_true',
        help=(
            'Share one y-axis across all plots; max derived from the batch '
            'largest peak (top major tick e.g. 300 when peak≈301)'
        ),
    )
    parser.add_argument(
        '--ylim-max',
        type=float,
        default=None,
        help='Force shared y-axis upper limit (implies standardized y)',
    )
    parser.add_argument(
        '--combine-grid',
        action='store_true',
        help=(
            'Also write a composite PNG: columns=persons, rows='
            'En Bloc → Fragmentary → No Memory Loss'
        ),
    )
    parser.add_argument(
        '--combine-output',
        type=Path,
        default=None,
        help='Path for combined grid PNG (default: <output-dir>/manuscript_curve_grid.png)',
    )
    # Labeling
    parser.add_argument(
        '--label-preset',
        choices=sorted(LABEL_PRESETS.keys()),
        default=None,
        help='Builtin value→title map (e.g. memlev_day)',
    )
    parser.add_argument(
        '--label-table',
        type=Path,
        default=None,
        help='Excel/CSV providing label values keyed by subid+curve_id',
    )
    parser.add_argument(
        '--label-sheet',
        default=None,
        help='Sheet name for --label-table Excel (default: Selected or Curves)',
    )
    parser.add_argument(
        '--label-col',
        default='memlev',
        help='Column in label table with raw label values (default memlev)',
    )
    parser.add_argument(
        '--label-map',
        default=None,
        help='JSON object/file mapping raw values→titles (or preset name)',
    )
    parser.add_argument(
        '--label-pair',
        nargs='+',
        default=None,
        help='Explicit titles: subid:curve_id:Title (spaces allowed in Title)',
    )
    parser.add_argument(
        '--label-config-json',
        type=Path,
        default=None,
        help='Full label_config dict as JSON file (overrides table/preset flags)',
    )
    args = parser.parse_args(argv)
    if not args.pairs and not args.from_selected and not args.label_pair:
        parser.error('Provide --pairs, --from-selected, and/or --label-pair')
    return args


def _label_config_from_args(args: argparse.Namespace) -> Optional[Dict[str, Any]]:
    if args.label_config_json is not None:
        path = (
            args.label_config_json
            if args.label_config_json.is_absolute()
            else PROJECT_ROOT / args.label_config_json
        )
        with open(path, 'r') as f:
            cfg = json.load(f)
        if not isinstance(cfg, dict):
            raise ValueError('label-config-json must contain a JSON object')
        return cfg

    # Auto table from --from-selected when using memlev preset / label-col
    table_path = args.label_table
    if table_path is None and args.from_selected is not None and (
        args.label_preset or args.label_map or args.label_col
    ):
        table_path = args.from_selected

    if table_path is None and not args.label_pair:
        return None

    cfg: Optional[Dict[str, Any]] = None
    if table_path is not None:
        path = table_path if table_path.is_absolute() else PROJECT_ROOT / table_path
        map_spec: Any = args.label_map or args.label_preset or 'memlev_day'
        cfg = {
            'source': 'table',
            'path': str(path),
            'sheet': args.label_sheet,
            'key_cols': ['subid', 'curve_id'],
            'value_col': args.label_col,
            'map': map_spec,
            'default': '',
        }
        # Only filter when explicitly using Curves (Selected is already curated)
        if args.label_sheet == 'Curves':
            cfg['filter_col'] = 'selected'
            cfg['filter_value'] = 1
    return cfg


def main(argv: Optional[Sequence[str]] = None) -> None:
    args = parse_args(argv)
    pairs: List[Pair] = []
    explicit_labels: LabelLookup = {}

    if args.label_pair:
        explicit_labels = parse_label_pairs(args.label_pair)
        pairs.extend(explicit_labels.keys())

    if args.from_selected is not None:
        sel_path = (
            args.from_selected
            if args.from_selected.is_absolute()
            else PROJECT_ROOT / args.from_selected
        )
        pairs.extend(pairs_from_selected_workbook(sel_path))

    if args.pairs:
        pairs.extend(parse_pairs(args.pairs))

    seen = set()
    unique_pairs: List[Pair] = []
    for p in pairs:
        if p not in seen:
            seen.add(p)
            unique_pairs.append(p)

    label_config = _label_config_from_args(args)

    manifest = export_pairs(
        cohort=args.cohort,
        pairs=unique_pairs,
        processed_subdir=args.processed_subdir,
        dataset_id=args.dataset_id,
        output_dir=args.output_dir,
        figsize=(float(args.figsize[0]), float(args.figsize[1])),
        show_peak=not args.no_peak,
        label_config=label_config,
        label_lookup=explicit_labels or None,
        standardize_y=bool(args.standardize_y or args.ylim_max is not None),
        ylim_max=args.ylim_max,
        style=args.style,
    )

    if args.combine_grid and manifest is not None and not manifest.empty:
        out_dir = resolve_output_dir(
            args.cohort, args.output_dir, PROJECT_ROOT, style=args.style,
        )
        grid_path = args.combine_output
        if grid_path is None:
            grid_path = out_dir / 'manuscript_curve_grid.png'
        elif not grid_path.is_absolute():
            grid_path = PROJECT_ROOT / grid_path
        combine_manuscript_grid(manifest, grid_path)


if __name__ == '__main__':
    main()
