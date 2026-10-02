from tokenize import group
import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.dates as mdates
import matplotlib
matplotlib.use("Agg")
import os
from ..Configuration.configuration import normalize_column
from ..Machine_Learning.get_feature_importances import get_feature_importances
from .plotting_utils import *
import numpy as np
from sklearn.tree import plot_tree


def _resolve_fontfamily(requested):
  """Prefer ``requested``; fall back to common Times-compatible faces on Linux."""
  if not requested:
    return None
  try:
    from matplotlib import font_manager
    available = {f.name for f in font_manager.fontManager.ttflist}
  except Exception:
    available = set()
  for name in (requested, 'Times New Roman', 'Nimbus Roman', 'Liberation Serif',
               'STIXGeneral', 'DejaVu Serif'):
    if name in available:
      return name
  return requested


_DEFAULT_TITLE_STATS_BBOX = {
  'boxstyle': 'round,pad=0.35',
  'facecolor': 'white',
  'edgecolor': 'black',
  'linewidth': 0.9,
  'alpha': 1.0,
}

def plot_smoothed_curve(df, plot_path, subid, dataset_identifier, event_number, peak, curve_threshold, title = "TAC Curve", event_timestamps = {}, df_version = 'SEARCH', subtitle_text = '', tac_column = 'TAC'):
  peak_time = df.loc[df[tac_column]==peak, 'datetime']
  # graph_cutoff = curve_ends + ((len(df) - curve_ends)*0.25)
  # df = df.loc[:graph_cutoff]
  
  fig, ax = plt.subplots(figsize = (16, 7))
  ax.plot(df['datetime'], df[tac_column], c='black')

  #annotate with lines for curve threshold, peak, curve_begin, curve_end
  ax.vlines(peak_time, ymin=curve_threshold, ymax=peak, color='black', linestyle='--')
  ax.hlines(curve_threshold, xmin=df['datetime'].min(), xmax=df['datetime'].max(), colors='black', linestyle='--')
  # ax.vlines(
  #   [df.loc[curve_begins, 'datetime'], df.loc[curve_ends, 'datetime']], 
  #   ymin=0, ymax=df['TAC'].max(), color='red', linestyle='--'
  # )

  ax.set_xlabel('Time (hrs)', fontsize = 24)
  ax.set_ylabel('TAC' if 'TAC' in tac_column else 'TAC (Pre-Imputation)', fontsize = 24)
  ax.tick_params(axis='x', labelsize = 18)
  ax.tick_params(axis='y', labelsize = 20)

  ax.set_title(title, fontsize=32, fontweight="semibold", pad=25)
  ax.text(0.5, 1.025, subtitle_text, fontsize=12, style='italic',
    ha='center', va='center', transform=ax.transAxes)
  # if drink_total is not None:
  #   ax.text(0.5, 1.05, f"Drink Total: {drink_total}", fontsize=14, ha='center', va='top', transform=ax.transAxes)

  ax.xaxis.set_major_locator(mdates.HourLocator(interval=2))
  ax.xaxis.set_major_formatter(mdates.DateFormatter('%H:%M'))

  if ax.get_ylim()[1] < 10:
    ax.set_ylim(-10, 20)

  if event_timestamps and all(value is not None for value in event_timestamps.values()):
    plot_event_lines(df, ax, event_timestamps, 'datetime', 'datetime')
  df_version = df_version if tac_column == 'TAC' else f'{df_version}_raw'
  path = f'{plot_path}{subid}_{dataset_identifier}_{event_number}_TAC_curve_{df_version}.png'
  plt.tight_layout()
  plt.savefig(path, bbox_inches='tight')
  plt.close('all')
  return path

def plot_signal_processing(df, plot_path, subid, event_number, dataset_identifier, df_version, 
                           curve_threshold, time_variable='datetime', title='Signal Processing', 
                           event_timestamps={}, subtitle_text='', show_imputations = True,
                           clean_event_labels=False, title_stats='',
                           axis_label_fontsize=None, tick_label_fontsize=None,
                           x_tick_label_fontsize=None, annotation_fontsize=10,
                           title_fontsize=18, title_stats_fontsize=11,
                           legend_fontsize=None, axis_label_fontweight=None,
                           time_tick_format='%H:%M', x_tick_rotation=45,
                           legend_bbox_to_anchor=(1.1, 1.1),
                           title_ha='center', title_x=0.5,
                           show_legend=True, filename_suffix='',
                           figsize=(16, 7), flagged_marker_size=20,
                           passed_marker_size=36, imputed_marker_size=36,
                           xlabel='Time', dpi=None,
                           title_stats_ha=None, title_stats_x=None,
                           title_stats_same_line=False,
                           fontfamily=None, title_stats_bbox=None,
                           pin_title=False):
  """
  Plot TAC signal-processing figure.

  clean_event_labels : bool
      If True, rename drinkStart_* / drinkFinish_* keys to "Drink Start" / "Drink End"
      before drawing. Default False preserves raw pipeline labels (skyn_dataset).
  title_stats : str
      Optional Drinks/eBAC line under the main title (e.g. "Drinks: 6  |  eBAC: 0.085").
  subtitle_text : str
      Optional middle subtitle (e.g. flagging reasons). With ``pin_title=True`` or
      non-empty ``title_stats``, drawn in the pinned manuscript header; otherwise
      uses the classic centered italic subtitle under ``set_title``.
  pin_title : bool
      If True, always use the pinned manuscript title layout (title / optional
      subtitle / optional title_stats) after tight_layout.
  """
  resolved_font = _resolve_fontfamily(fontfamily)
  _prev_font = None
  if resolved_font:
    _prev_font = plt.rcParams['font.family']
    plt.rcParams['font.family'] = resolved_font
  try:
    return _plot_signal_processing_body(
      df, plot_path, subid, event_number, dataset_identifier, df_version,
      curve_threshold, time_variable=time_variable, title=title,
      event_timestamps=event_timestamps, subtitle_text=subtitle_text,
      show_imputations=show_imputations, clean_event_labels=clean_event_labels,
      title_stats=title_stats, axis_label_fontsize=axis_label_fontsize,
      tick_label_fontsize=tick_label_fontsize,
      x_tick_label_fontsize=x_tick_label_fontsize,
      annotation_fontsize=annotation_fontsize, title_fontsize=title_fontsize,
      title_stats_fontsize=title_stats_fontsize, legend_fontsize=legend_fontsize,
      axis_label_fontweight=axis_label_fontweight,
      time_tick_format=time_tick_format, x_tick_rotation=x_tick_rotation,
      legend_bbox_to_anchor=legend_bbox_to_anchor, title_ha=title_ha,
      title_x=title_x, show_legend=show_legend, filename_suffix=filename_suffix,
      figsize=figsize, flagged_marker_size=flagged_marker_size,
      passed_marker_size=passed_marker_size,
      imputed_marker_size=imputed_marker_size, xlabel=xlabel, dpi=dpi,
      title_stats_ha=title_stats_ha, title_stats_x=title_stats_x,
      title_stats_same_line=title_stats_same_line,
      title_stats_bbox=title_stats_bbox, pin_title=pin_title,
    )
  finally:
    if _prev_font is not None:
      plt.rcParams['font.family'] = _prev_font


def _plot_signal_processing_body(df, plot_path, subid, event_number, dataset_identifier, df_version, 
                           curve_threshold, time_variable='datetime', title='Signal Processing', 
                           event_timestamps={}, subtitle_text='', show_imputations = True,
                           clean_event_labels=False, title_stats='',
                           axis_label_fontsize=None, tick_label_fontsize=None,
                           x_tick_label_fontsize=None, annotation_fontsize=10,
                           title_fontsize=18, title_stats_fontsize=11,
                           legend_fontsize=None, axis_label_fontweight=None,
                           time_tick_format='%H:%M', x_tick_rotation=45,
                           legend_bbox_to_anchor=(1.1, 1.1),
                           title_ha='center', title_x=0.5,
                           show_legend=True, filename_suffix='',
                           figsize=(16, 7), flagged_marker_size=20,
                           passed_marker_size=36, imputed_marker_size=36,
                           xlabel='Time', dpi=None,
                           title_stats_ha=None, title_stats_x=None,
                           title_stats_same_line=False,
                           title_stats_bbox=None, pin_title=False):
  """Inner implementation of ``plot_signal_processing`` (font already applied)."""
  passed = df.loc[
    (df['imp_cand']==0)
  ]

  # For showing non-imputed data
  gap = df.loc[df['gap'] == 1]
  non_wear = df.loc[(df['non_wear'] == 1)]
  jumps = df.loc[(df['jump_imp_cand'] == 1)]
  plummet = df.loc[(df['plummet_imp_cand'] == 1)]
  extreme_negative = df.loc[(df['extreme_negative_imp_cand'] == 1)]
  proximal_low_quality = df.loc[(df['proximal_low_quality_imp_cand'] == 1)]

  # Create a figure and axis
  fig, ax = plt.subplots(figsize=figsize)
  
  #Smoothed Final TAC
  ax.plot(df[time_variable], df['TAC' if show_imputations else 'TAC_pre_imputation'], label="TAC (Processed)", alpha=0.5, color="black", linewidth = 2)
  
  #Passed (high quality values) - use final TAC since passed data is not imputed
  ax.scatter(passed[time_variable], passed['TAC' if show_imputations else 'TAC_pre_imputation'], label='Passed', 
             color='darkblue', marker='.', alpha=1.0, s=passed_marker_size)
  #Non Wear
  if not non_wear.empty:
    ax.scatter(non_wear[time_variable], non_wear['TAC_pre_imputation'], label='Non-Wear', 
             color='lightpink', marker='x', alpha=0.7, s=flagged_marker_size)
  #Extreme Negative (display label: Very Negative)
  if not extreme_negative.empty:
    ax.scatter(extreme_negative[time_variable], extreme_negative['TAC_pre_imputation'], label='Very Negative', 
             color='lightsteelblue', marker='*', alpha=0.7, s=flagged_marker_size)
  #Jumps
  if not jumps.empty:
    ax.scatter(jumps[time_variable], jumps['TAC_pre_imputation'], label='Jump', 
              color='lightblue', marker='^', alpha=0.7, s=flagged_marker_size)
  #Plummet
  if not plummet.empty:
    ax.scatter(plummet[time_variable], plummet['TAC_pre_imputation'], label='Plummet', 
              color='thistle', marker='v', alpha=0.7, s=flagged_marker_size)
  # Between low quality
  if not proximal_low_quality.empty:
    ax.scatter(proximal_low_quality[time_variable], proximal_low_quality['TAC_pre_imputation'],
               label='Proximal Low Quality', color='gray', marker='s', alpha=0.7, s=flagged_marker_size)
    
  # Imputed data
  if show_imputations:
    gap_imputed = df.loc[df['gap_imputed'] == 1]
    non_wear_imputed = df.loc[(df['non_wear_imputed'] == 1)]
    extreme_negative_imputed = df.loc[(df['extreme_negative_imputed'] == 1)]
    jump_imputed = df.loc[df['jump_imputed'] == 1]
    plummet_imputed = df.loc[df['plummet_imputed'] == 1]
    proximal_low_quality_imputed = df.loc[df['proximal_low_quality_imputed'] == 1]
    if not gap_imputed.empty:
      ax.scatter(gap_imputed[time_variable], gap_imputed['TAC_pre_savgol'], label='Imputed Gap', 
                marker='o', alpha=1.0, facecolor='gray', edgecolors="black", s=imputed_marker_size)
    if not non_wear_imputed.empty:
      ax.scatter(non_wear_imputed[time_variable], non_wear_imputed['TAC_pre_savgol'], 
                label='Imputed Non-Wear', facecolor='lightpink', edgecolors= "darkred", marker='o', alpha=1.0,
                s=imputed_marker_size)
    if not extreme_negative_imputed.empty:
      ax.scatter(extreme_negative_imputed[time_variable], extreme_negative_imputed['TAC_pre_savgol'], 
                label='Imputed Very Negative', facecolor='lightsteelblue', edgecolors= "purple", marker='o', alpha=1.0,
                s=imputed_marker_size)
    if not jump_imputed.empty:
      ax.scatter(jump_imputed[time_variable], jump_imputed['TAC_pre_savgol'], 
                label='Imputed Jump', facecolor='lightblue', edgecolors= "darkblue", marker='o', alpha=1.0,
                s=imputed_marker_size)
    if not plummet_imputed.empty:
      ax.scatter(plummet_imputed[time_variable], plummet_imputed['TAC_pre_savgol'], 
                label='Imputed Plummet', facecolor='thistle', edgecolors= "purple", marker='o', alpha=1.0,
                s=imputed_marker_size)
    if not proximal_low_quality_imputed.empty:
      ax.scatter(proximal_low_quality_imputed[time_variable], proximal_low_quality_imputed['TAC_pre_savgol'],
                label='Imputed Proximal Low Quality', facecolor='gray', edgecolors="darkgreen", marker='o', alpha=1.0,
                s=imputed_marker_size)
  
  # Plot threshold line
  ax.hlines(curve_threshold, xmin=df['datetime'].min(), xmax=df['datetime'].max(), 
            colors='black', linestyle='--', label="Curve Threshold")

  # Plot event timestamps if available
  plot_events = event_timestamps
  if clean_event_labels and event_timestamps:
    plot_events = {}
    for key, ts in event_timestamps.items():
      k = str(key)
      if 'drinkStart' in k:
        label = 'Drink Start'
      elif 'drinkFinish' in k:
        label = 'Drink End'
      else:
        label = key
      # Avoid colliding keys if multiple drink events share a cleaned name
      if label in plot_events and label != key:
        n = 2
        while f'{label} ({n})' in plot_events:
          n += 1
        label = f'{label} ({n})'
      plot_events[label] = ts

  if plot_events and all(value is not None for value in plot_events.values()):
    plot_event_lines(
      df, ax, plot_events, 'datetime', 'datetime', font_size=annotation_fontsize,
    )

  # Format the x-axis for time
  ax.xaxis.set_major_locator(mdates.HourLocator(interval=2))
  ax.xaxis.set_major_formatter(mdates.DateFormatter(time_tick_format))

  # Add labels, title, and legend
  axis_kwargs = {}
  if axis_label_fontsize:
    axis_kwargs['fontsize'] = axis_label_fontsize
  if axis_label_fontweight:
    axis_kwargs['fontweight'] = axis_label_fontweight
  ax.set_xlabel(xlabel, **axis_kwargs)
  ax.set_ylabel('TAC', **axis_kwargs)
  x_tick_fs = x_tick_label_fontsize if x_tick_label_fontsize is not None else tick_label_fontsize
  y_tick_fs = tick_label_fontsize
  stats = (title_stats or '').strip()
  sub = (subtitle_text or '').strip()
  use_pinned = bool(pin_title) or bool(stats)
  if sub and not use_pinned:
    # Classic skyn_dataset header (title + centered italic subtitle)
    ax.set_title(title, fontsize=title_fontsize, fontweight="semibold", pad=25)
  for label in ax.get_xticklabels():
    label.set_rotation(x_tick_rotation)
    if x_tick_fs:
      label.set_fontsize(x_tick_fs)
  if y_tick_fs:
    for label in ax.get_yticklabels():
      label.set_fontsize(y_tick_fs)
  if show_legend:
    legend_kwargs = {}
    if legend_fontsize:
      legend_kwargs['fontsize'] = legend_fontsize
    ax.legend(loc='upper right', bbox_to_anchor=legend_bbox_to_anchor, ncol=2,
             frameon=True, framealpha=1, edgecolor='black', facecolor='white',
             **legend_kwargs)

  # Classic centered italic subtitle (only when not using pinned manuscript header)
  if sub and not use_pinned:
    ax.text(0.5, 1.025, sub, fontsize=10, style='italic',
            ha='center', va='center', transform=ax.transAxes)
  
  # Save the figure
  df_version = df_version if show_imputations else f'{df_version}_raw'
  suffix = filename_suffix or ''
  path = (
    f'{plot_path}{subid}_{dataset_identifier}_{event_number}'
    f'_TAC_processing_{df_version}{suffix}.png'
  )
  plt.tight_layout()
  # Manuscript-style title / optional flag subtitle / optional Drinks/eBAC
  if use_pinned:
    stats_ha = title_ha if title_stats_ha is None else title_stats_ha
    stats_x = title_x if title_stats_x is None else title_stats_x
    if title_stats_bbox is True:
      stats_bbox = dict(_DEFAULT_TITLE_STATS_BBOX)
    elif isinstance(title_stats_bbox, dict):
      stats_bbox = dict(_DEFAULT_TITLE_STATS_BBOX)
      stats_bbox.update(title_stats_bbox)
    else:
      stats_bbox = None
    stats_kwargs = {}
    if stats_bbox is not None:
      stats_kwargs['bbox'] = stats_bbox

    if title_stats_same_line and stats and not sub:
      y = 1.04
      ax.text(
        title_x, y, title, transform=ax.transAxes, fontsize=title_fontsize,
        fontweight='semibold', ha=title_ha, va='bottom', clip_on=False,
      )
      ax.text(
        stats_x, y, stats, transform=ax.transAxes, fontsize=title_stats_fontsize,
        ha=stats_ha, va='bottom', clip_on=False, **stats_kwargs,
      )
    elif sub and stats:
      ax.text(
        title_x, 1.14, title, transform=ax.transAxes, fontsize=title_fontsize,
        fontweight='semibold', ha=title_ha, va='bottom', clip_on=False,
      )
      ax.text(
        title_x, 1.07, sub, transform=ax.transAxes,
        fontsize=max(10, int(title_stats_fontsize) - 2),
        style='italic', ha=title_ha, va='bottom', clip_on=False,
      )
      ax.text(
        stats_x, 1.01, stats, transform=ax.transAxes, fontsize=title_stats_fontsize,
        ha=stats_ha, va='bottom', clip_on=False, **stats_kwargs,
      )
    elif stats:
      ax.text(
        title_x, 1.08, title, transform=ax.transAxes, fontsize=title_fontsize,
        fontweight='semibold', ha=title_ha, va='bottom', clip_on=False,
      )
      ax.text(
        stats_x, 1.015, stats, transform=ax.transAxes, fontsize=title_stats_fontsize,
        ha=stats_ha, va='bottom', clip_on=False, **stats_kwargs,
      )
    elif sub:
      ax.text(
        title_x, 1.08, title, transform=ax.transAxes, fontsize=title_fontsize,
        fontweight='semibold', ha=title_ha, va='bottom', clip_on=False,
      )
      ax.text(
        title_x, 1.015, sub, transform=ax.transAxes,
        fontsize=max(10, int(title_stats_fontsize) - 2),
        style='italic', ha=title_ha, va='bottom', clip_on=False,
      )
    else:
      ax.text(
        title_x, 1.015, title, transform=ax.transAxes, fontsize=title_fontsize,
        fontweight='semibold', ha=title_ha, va='bottom', clip_on=False,
      )
  save_kwargs = {'bbox_inches': 'tight'}
  if dpi is not None:
    save_kwargs['dpi'] = dpi
  plt.savefig(path, **save_kwargs)
  plt.close('all')
  
  return path

def plot_tac_and_temp(df, plot_folder, subid, dataset_identifier, event_number, tac_variable, temp_variable, time_variable, plot_title = "TAC and Temperature", event_timestamps = {}, df_version = 'SEARCH', drink_total = None):
  fig, ax1 = plt.subplots(figsize=(16, 7))
  ax1.plot(df[time_variable], df[tac_variable], color="darkblue", label=tac_variable)
  ax2 = ax1.twinx()
  ax2.plot(df[time_variable], df[temp_variable], color="maroon", label=temp_variable)
  
  ax1.set_title(plot_title, fontsize=26, fontweight="semibold", pad=25 if drink_total is not None else 15)
  if drink_total is not None:
    ax1.text(0.5, 1.05, f"Drink Total: {drink_total}", fontsize=14, ha='center', va='top', transform=ax1.transAxes)
  
  ax1.set_ylabel(tac_variable, color='darkblue')
  ax2.set_ylabel(temp_variable, color='maroon')
  
  ax1.legend(loc='upper left')
  ax2.legend(loc='upper right')
  
  if df[tac_variable].max() < 10:
    ax1.set_ylim(0, 10)
  if df[temp_variable].min() > 15:
    ax2.set_ylim(15, 40) 
    
  if event_timestamps and all(value is not None for value in event_timestamps.values()):
    plot_event_lines(df, ax1, event_timestamps, time_variable, 'datetime')
  
  ax1.xaxis.set_major_locator(mdates.HourLocator(interval=2))
  ax1.xaxis.set_major_formatter(mdates.DateFormatter('%H:%M'))

  path = f'{plot_folder}tac_and_temp_plot_{subid}_{dataset_identifier}_{event_number}_{df_version}.png'
  fig.savefig(path)
  plt.close('all')

  return path

def plot_column(df, plot_folder, subid, condition, dataset_identifier, y_variable, time_variable, xlabel="Time (hours)", event_timestamps = {}):
      ylabel = y_variable    
      title = f'{y_variable} {subid} {condition}'
      fig, ax = plt.subplots(figsize=(16, 7)) 
      ax.scatter(y=df[y_variable], x=df[time_variable])
      ax.set_title(title, fontsize=26, fontweight="semibold", pad=15)
      ax.set_ylabel(ylabel)
      ax.set_xlabel(xlabel)
      plot_event_lines(df, ax, event_timestamps, time_variable, 'datetime')
      full_path = f'{plot_folder}/{y_variable} - {subid} - {condition}{dataset_identifier}.png'
      fig.savefig(full_path)
      plt.close('all')
      return full_path

def plot_TAC_curve(df, plot_folder, subid, condition, dataset_identifier, tac_variable, time_variable, ylabel="TAC ug/L", xlabel="Time (hours)", event_timestamps = {}):
  title = f'TAC Curve - {tac_variable}'
  fig, ax = plt.subplots(figsize=(16, 7)) 
  ax.scatter(df[time_variable], df[tac_variable])
  ax.set_title(title, fontsize=26, fontweight="semibold", pad=15)
  ax.set_ylabel(ylabel)
  ax.set_xlabel(xlabel)
  if df[tac_variable].max() < 15:
        ax.set_ylim(df[tac_variable].min() - 1, 15)
  plot_event_lines(df, ax, event_timestamps, time_variable, 'datetime')

  full_path = f'{plot_folder}/{tac_variable} - {subid} - {condition}{dataset_identifier}.png'
  fig.savefig(full_path)
  plt.close('all')
  return full_path

def plot_overlaid_TAC_curves(df, plot_folder, subid, condition, dataset_identifier, tac_variables, time_variable, plot_name, ylabel="TAC ug/L", xlabel="Time (hours)", event_timestamps = {}):
  title = f'TAC Curve - {plot_name}'
  fig, ax = plt.subplots(figsize=(16, 7))
  for variable in tac_variables:
    ax.plot(df[time_variable], df[variable])
  ax.set_title(title, fontsize=26, fontweight="semibold", pad=15)
  ax.set_ylabel(ylabel)
  ax.set_xlabel(xlabel)
  plot_event_lines(df, ax, event_timestamps, time_variable, 'datetime')
  if ax.get_ylim()[1] < 10:
    ax.set_ylim(0, 10)

  full_path = f'{plot_folder}/{plot_name} - {subid} - {condition}{dataset_identifier}.png'
  fig.savefig(full_path)
  plt.close('all')
  return full_path

def plot_cropping(data, beginning_timestamp, end_timestamp, self, add_color = False, title="Raw Data Cropping"):
      if add_color:
            marker_colors = ['red', 'black']
      else:
            marker_colors = ['lightgray', 'black']

      hours = (data['datetime'] - beginning_timestamp) / np.timedelta64(1, 'h')
      cropped = data.loc[(data['datetime'] < beginning_timestamp) | (data['datetime'] > end_timestamp), 'TAC']
      cropped_time = hours.loc[(data['datetime'] < beginning_timestamp) | (data['datetime'] > end_timestamp)]

      kept = data.loc[(data['datetime'] > beginning_timestamp) & (data['datetime'] < end_timestamp), 'TAC']
      kept_time = hours.loc[(data['datetime'] > beginning_timestamp) & (data['datetime'] < end_timestamp)]
      
      try:
            cropped_at_max_duration = round(kept_time.max()) == round(self.max_duration)
            timestamps_valid = True
      except:
            timestamps_valid = False
            title = "Invalid Timestamps - Cropping Failed"

      fig, ax = plt.subplots(figsize=(16, 7))
      ax.scatter(kept_time, kept, marker='.', c = marker_colors[1], s=12)      
      ax.scatter(cropped_time, cropped, marker='x', c = marker_colors[0], s=14)
      #ax.legend(("Kept", "Cropped"), loc='upper right', fontsize=16)
      ax.set_xlabel('Time (hrs)', fontsize = 24)
      ax.set_ylabel('TAC', fontsize = 24, labelpad=3)
      ax.set_title(title, fontsize = 36, fontweight="semibold", pad=15)
      if ax.get_ylim()[1] < 10:
            ax.set_ylim(0, 10)
      ax.tick_params(axis = 'x', labelsize = 22)
      ax.tick_params(axis = 'y', labelsize = 22)
      if timestamps_valid:
            if beginning_timestamp > data['datetime'].min():
                  ax.vlines(x = 0, ymin = -8, ymax = data['TAC'].max(), color = 'black', linestyle = '--')
                  plt.text(0.2, data['TAC'].max() * 0.9, "Episode", fontsize = 28, fontstyle = "italic")
                  plt.text(0.2, data['TAC'].max() * 0.8, "Beginning", fontsize = 28, fontstyle = "italic")
            if end_timestamp < data['datetime'].max():
                  ax.vlines(x = kept_time.max(), ymin = -8, ymax = data['TAC'].max(), color = 'black', linestyle = '--')
                  cutoff_label = f'{round(self.max_duration)}-Hr' if cropped_at_max_duration else 'Episode End'
                  plt.text(kept_time.max()*0.85, data['TAC'].max() * 0.9, cutoff_label, fontsize = 28, fontstyle = "italic")
                  if cropped_at_max_duration:
                        plt.text(kept_time.max()*0.85, data['TAC'].max() * 0.8, "Cutoff", fontsize = 28, fontstyle = "italic")
      path = f'{self.plot_folder}/{self.subid}_{self.condition}{self.dataset_identifier}_cropping.png'
      plt.tight_layout()
      plt.savefig(path, bbox_inches='tight')
      plt.close('all')
      return path

def plot_cleaning_comparison(occasion, df, time_variable, add_color = False, title = "TAC Signal Cleaning", size = (16, 7), legend = True, snip = None, event_timestamps = {}):

      if snip:
            df = df.iloc[snip[0]:snip[1]]

      tac_columns = ['TAC_Raw', 'TAC_sloped_start_reassigned', 'TAC_negative_reassigned', 'TAC_gaps_filled', 'TAC_extreme_values_imputed', 'TAC_device_off', 'TAC_processed']
      markers = [',', 'o', 'd', 'H', 'v', 's', 'p']
      names = ['Raw (passed)', 'Sloped Beginning Reassigned', 'Subzero Value Reassigned', 'Imputed (Missing Value)', 'Imputed (Extreme Outlier)', 'Imputed (Device Off)', 'Imputed (Artifact)']

      marker_colors = ['dimgray', 'yellow', 'steelblue', 'peru', 'firebrick', 'mediumvioletred', 'maroon'] if add_color else ['black', 'dimgray', 'dimgray', 'dimgray', 'dimgray', 'dimgray', 'dimgray']

      all_imputed_idx = []
      
      tac_plot_specs = {}
      for i, col in enumerate(tac_columns):
            if col in df.columns.tolist():
                  if i > 0 and (i < len(tac_columns) - 1):
                        df_current_col = df[col]
                        df_prior_cols = df[[col for col in tac_columns[:i] if col in df.columns]]
                        df_post_cols = df[[col for col in tac_columns[i+1:] if col in df.columns]]
                        
                        mask_different = df_prior_cols.ne(df_current_col, axis='index').all(axis=1)
                        mask_equal = df_post_cols.eq(df_current_col, axis='index').all(axis=1)

                        combined_mask = mask_different & mask_equal
                        imputed_idx = combined_mask[combined_mask].index
                        all_imputed_idx.extend(imputed_idx)

                  elif i == len(tac_columns) -1:
                        df_current_col = df[col]
                        df_prior_cols = df[[col for col in tac_columns[:i] if col in df.columns]]

                        mask_different = df_prior_cols.ne(df_current_col, axis='index').all(axis=1)
                        imputed_idx = mask_different[mask_different].index
                        all_imputed_idx.extend(imputed_idx)

                  else:
                        imputed_idx = []

                  if len(imputed_idx) > 0:
                        tac_plot_specs[names[i]] = {
                              'x': df.loc[imputed_idx, time_variable].tolist(),
                              'y': df.loc[imputed_idx, col].tolist(),
                              'color': marker_colors[i],
                              'marker': markers[i],
                        }
            
      tac_plot_specs['Raw (passed)'] = {
            'x': df.loc[[i for i in df.index.tolist() if i not in all_imputed_idx], time_variable].tolist(),
            'y': df.loc[[i for i in df.index.tolist() if i not in all_imputed_idx], 'TAC_Raw'].tolist(),
            'color': marker_colors[0],
            'marker': markers[0]
      }

      tac_plot_specs['Raw (removed)'] = {
            'x': df.loc[all_imputed_idx, time_variable].tolist(),
            'y': df.loc[all_imputed_idx, 'TAC_Raw'].tolist(),
            'color': 'crimson' if add_color else 'darkgrey',
            'marker': 'x'
      }

      smoothed_tac = df.loc[:, f'TAC_processed_smooth_{occasion.smoothing_window}'].tolist()
      smoothed_tac_x = df.loc[:, time_variable].tolist()
      fig, ax = plt.subplots(figsize=size)
      ax.plot(smoothed_tac_x, smoothed_tac, linestyle = 'solid', color = 'midnightblue' if add_color else 'dimgray', linewidth = 2)
      for name, specs in tac_plot_specs.items():
            ax.scatter(specs['x'], specs['y'], c = specs['color'], marker = specs['marker'], s = 20)
            if add_color and len(specs['x']) and name not in ['Raw (passed)', 'Raw (removed)']:
                  ranges = find_ranges(specs['x'], threshold=0.05)
                  for start, end in ranges:
                        ax.axvspan(start, end, color=specs['color'], alpha=0.2, label='_nolegend_')

      ax.set_xlabel('Time (hrs)', fontsize = 22)
      ax.set_ylabel('TAC', fontsize = 22)
      if ax.get_ylim()[1] < 10:
            ax.set_ylim(0, 10)
      ax.set_title(title, fontsize=28, fontweight="semibold", pad=15)
      if not occasion.valid_occasion:
            split_text_index = int(round(len(occasion.invalid_reason)/2))
            split_character_is_space = occasion.invalid_reason[split_text_index] == ' '
            while not split_character_is_space:
                  split_text_index += 1
                  split_character_is_space = occasion.invalid_reason[split_text_index] == ' '
                  if split_character_is_space or split_text_index == len(occasion.invalid_reason) - 1:
                        break
            invalid_reason_part_1 = occasion.invalid_reason[:split_text_index]
            invalid_reason_part_2 = occasion.invalid_reason[split_text_index:]
            plt.text(0.03, 0.95, f'Invalid: {invalid_reason_part_1}', ha='left', va='center', transform=plt.gca().transAxes, fontdict={'size':  9})
            plt.text(0.03, 0.92, f'{invalid_reason_part_2}', ha='left', va='center', transform=plt.gca().transAxes, fontdict={'size':  9})
      if legend:
            ax.legend(["Smooth (Final)"] + list(tac_plot_specs.keys()), loc='upper right', fontsize=16)
      ax.tick_params(axis='x', labelsize = 18)
      ax.tick_params(axis='y', labelsize = 18)

      plot_event_lines(df, ax, event_timestamps, time_variable, 'datetime')
      
      filename = f'{occasion.plot_folder}cleaning - {occasion.subid} - {occasion.condition}{occasion.dataset_identifier}'
      if snip:
            filename = filename + f' - snip{snip[0]}_{snip[1]}'
      filename = filename + '.png'
      plt.tight_layout()
      plt.savefig(filename, bbox_inches='tight')
      plt.close('all')
      return filename

def create_temp_histogram(self):
  temperatures = []
  for occasion in self.occasions:
    data = occasion.dataset['Temperature_C'].tolist()
    temperatures.extend(data)
  temperatures_below_threshold = [temp for temp in temperatures if temp < 27]

  plt.hist(temperatures, bins = [20, 22, 24, 26, 28, 30, 32, 34, 36, 38, 40])
  plt.ylabel("Counts (1000's)")
  plt.xlabel('Temperature (Celsius)')
  plt.title('Temperature Distribution')
  plt.xticks([20, 22, 24, 26, 28, 30, 32, 34, 36, 38, 40])
  plt.yticks(ticks = [2000, 4000, 6000, 8000, 10000, 12000, 14000, 16000, 18000], labels=[2, 4, 6, 8, 10, 12, 14, 16, 18])
  plt.savefig(f'{self.analyses_out_folder}/temperature_histogram.png')
  plt.close('all')


def create_simple_histogram_of_feature(feature, save_folder, feature_name):
  # feature_names = [
  #         'ending_non_wear_perc_CURVE', 'flatline_max_SEARCH', 'flatline_max_CURVE',
  #         'device_turned_on_percent_CURVE', 'device_turned_on_duration_CURVE', 'device_turned_on_percent_CURVE', 
  #         'device_worn_duration_CURVE', 'device_worn_percent_CURVE', 'device_worn_percent_of_device_on_CURVE',
  #         'negative_duration_CURVE', 'sub_negative_10_duration_CURVE', 'duration_CURVE', 'first_tac_CURVE', 'last_tac_CURVE',
  #         'mean_tac_CURVE', 'peak_CURVE', 'auc_total_CURVE', 'rise_duration_CURVE', 'fall_duration_CURVE', 'rise_rate_CURVE', 
  #         'fall_rate_CURVE', 'fall_complete_percent_CURVE', ''
  #       ]
  #       for feature in feature_names:
  #         save_feature
  #         if not os.path.exists(processed_data_out):
  #         create_simple_histogram_of_feature(self.event_level_data[feature], f'{self.data_out_folder}/)

  plt.hist(feature, bins = [20, 22, 24, 26, 28, 30, 32, 34, 36, 38, 40])
  plt.xlabel(feature_name)
  plt.title(f'{feature_name} Distribution')
  plt.savefig(f'{save_folder}/{feature_name}_histogram.png')
  plt.close('all')

def plot_box_whisker(features, variables, ground_truth_variable, plot_folder, cohort_name, filter = {}):
      features = features[features['valid_occasion']==1].reset_index(drop=True)
      for variable_name, exclude_values in filter.items():
            features = features[~features[variable_name].isin(exclude_values)]
      data = {}
      for variable in variables:
            data[variable] = [[] for i in range(0, len(features[ground_truth_variable].unique()))]
            for i, truth in enumerate(features[ground_truth_variable].unique()):
                  data[variable][i] = features[features[ground_truth_variable]==truth][variable].tolist()
      
      figure, axis = plt.subplots(1, len(variables))
      figure.set_figheight(12)
      figure.set_figwidth(10 * len(variables))
      colors = ['firebrick', 'silver', 'darkblue']
      ticker = 0
      positions = [i+1 for i in range(0, len(features[ground_truth_variable].unique()))]
      for variable in variables:
            bp = axis[0 + ticker].boxplot(data[variable], positions = positions, widths = 0.5, patch_artist = True)
            axis[0 + ticker].set_title(variable)
            axis[0 + ticker].set_xticks([(i+1) for i in range(0, len(features[ground_truth_variable].unique()))])
            axis[0 + ticker].set_xticklabels(features[ground_truth_variable].unique().tolist())
            axis[0 + ticker].grid(color = 'lightgray', linestyle = '--', linewidth = 0.75, axis='y')
            ticker += 1
            for patch, color in zip(bp['boxes'], colors):
                  patch.set_facecolor(color)
            for median in bp['medians']:
                  median.set(color ='black', linewidth = 2)

            
      full_path = f'{plot_folder}/features_{cohort_name}_{ground_truth_variable}_box_whisker.png'
      figure.savefig(full_path, bbox_inches='tight')
      plt.close('all')


def plot_rf_feature_importances(model, model_name, model_figures_folder):
  forest_importance = get_feature_importances(model.best, model.predictors)
  fig, ax = plt.subplots()
  ax.grid(True, axis = 'x')
  forest_importance.plot.barh(ax=ax, color='dimgray')
  ax.set_title(f'Feature Importance - Random Forest', fontdict={'weight': 'bold', 'size': 15})
  ax.set_xlabel("Mean Decrease in Impurity", fontdict={'size': 13})
  ax.get_legend().remove()
  fig.tight_layout()

  plt.savefig(f'{model_figures_folder}/{model_name} - Feature Importances.png', dpi=55)
  plt.close('all')
  return forest_importance

def plot_rf_tree(rf, feature_names, cohort_name, model_figures_folder):
  fig = plt.figure(figsize=(15, 10))
  plot_tree(rf.best.estimators_[0], 
    feature_names=feature_names, class_names=['Alc', 'Non'], 
    filled=True, impurity=True, rounded=True)
  fig.savefig(f'{model_figures_folder}/Random Forest - {cohort_name} - Decision Tree.png')
  plt.close('all')


def plot_cluster_analysis(df, plot_folder, subid, dataset_identifier, curve_threshold, title="TAC Cluster Analysis", subtitle_text=""):
  """
  Visualize the different clusters and excluded data in the TAC data.
  
  Args:
    df (pd.DataFrame): Main dataframe with cluster labels
    plot_folder (str): Path to save the plot
    subid (str): Subject ID
    dataset_identifier (str): Dataset identifier
    curve_threshold (float): Curve threshold value
    title (str): Plot title
    subtitle_text (str): Subtitle text
  """
  # Check if cluster data exists
  if 'labeled_cluster_data' not in df.attrs:
    print("Warning: No cluster data found in dataframe attributes")
    return None
    
  labeled_data = df.attrs['labeled_cluster_data']
  
  # Create figure and axis
  fig, ax = plt.subplots(figsize=(16, 10))
  
  # Plot TAC data with different colors for each cluster
  colors = {
    'baseline': 'darkgreen',
    'excluded': 'lightgray'
  }
  
  # Generate colors for other clusters
  cluster_ids = labeled_data['cluster_id'].dropna().unique()
  if len(cluster_ids) > 0:
    import matplotlib.cm as cm
    cmap = cm.Set3
    for i, cluster_id in enumerate(cluster_ids):
      if cluster_id != labeled_data[labeled_data['is_baseline'] == True]['cluster_id'].iloc[0] if any(labeled_data['is_baseline']) else None:
        colors[f'cluster_{int(cluster_id)}'] = cmap(i % cmap.N)
  
  # Plot each cluster type
  for cluster_type in colors.keys():
    if cluster_type == 'excluded':
      # Plot excluded data
      excluded_mask = labeled_data['cluster_label'] == 'excluded'
      if excluded_mask.any():
        excluded_data = labeled_data[excluded_mask]
        ax.scatter(excluded_data['datetime'], excluded_data['TAC'], 
                  c=colors[cluster_type], alpha=0.3, s=10, 
                  label=f'Excluded ({excluded_mask.sum()} points)', marker='x')
    else:
      # Plot cluster data
      cluster_mask = labeled_data['cluster_label'] == cluster_type
      if cluster_mask.any():
        cluster_data = labeled_data[cluster_mask]
        ax.scatter(cluster_data['datetime'], cluster_data['TAC'], 
                  c=colors[cluster_type], alpha=0.7, s=15, 
                  label=f'{cluster_type.replace("_", " ").title()} ({cluster_mask.sum()} points)')
  
  # Add curve threshold line
  ax.axhline(y=curve_threshold, color='red', linestyle='--', linewidth=2, 
             label=f'Curve Threshold: {curve_threshold:.2f}')
  
  # Add zero line for reference
  ax.axhline(y=0, color='black', linestyle='-', linewidth=0.5, alpha=0.5)
  
  # Add TAC = -5 line to show exclusion boundary
  ax.axhline(y=-5, color='orange', linestyle=':', linewidth=1, alpha=0.7, 
             label='TAC = -5 (Exclusion Boundary)')
  
  # Format the plot
  ax.set_xlabel('Time', fontsize=14)
  ax.set_ylabel('TAC (μg/L)', fontsize=14)
  ax.set_title(title, fontsize=18, fontweight="semibold", pad=25)
  
  # Add subtitle
  if subtitle_text:
    ax.text(0.5, 1.025, subtitle_text, fontsize=12, style='italic',
            ha='center', va='center', transform=ax.transAxes)
  
  # Format x-axis for time
  ax.xaxis.set_major_locator(mdates.HourLocator(interval=2))
  ax.xaxis.set_major_formatter(mdates.DateFormatter('%H:%M'))
  
  # Add legend
  ax.legend(loc='upper right', bbox_to_anchor=(1.15, 1.0), 
           frameon=True, framealpha=1, edgecolor='black', facecolor='white')
  
  # Set y-axis limits to show full range
  y_min = labeled_data['TAC'].min()
  y_max = labeled_data['TAC'].max()
  y_range = y_max - y_min
  ax.set_ylim(y_min - 0.1 * y_range, y_max + 0.1 * y_range)
  
  # Add grid for better readability
  ax.grid(True, alpha=0.3)
  
  # Add exclusion reason summary in text box
  if 'exclusion_reason' in labeled_data.columns:
    exclusion_counts = labeled_data['exclusion_reason'].value_counts()
    exclusion_text = "Exclusion Reasons:\n"
    for reason, count in exclusion_counts.items():
      if reason != 'none':
        exclusion_text += f"• {reason}: {count}\n"
    
    if exclusion_text != "Exclusion Reasons:\n":
      ax.text(0.02, 0.98, exclusion_text, transform=ax.transAxes, 
              verticalalignment='top', fontsize=10, 
              bbox=dict(boxstyle='round', facecolor='white', alpha=0.8))
  
  # Save the plot
  path = f'{plot_folder}{subid}_{dataset_identifier}_cluster_analysis.png'
  plt.tight_layout()
  plt.savefig(path, bbox_inches='tight', dpi=300)
  plt.close('all')
  
  return path

