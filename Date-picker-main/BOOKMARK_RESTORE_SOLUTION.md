# Bookmark Restore Solution for Dynamic Presets

## Problem Summary

When a Power BI bookmark is created with a time-based preset (e.g., "Today") and attached to an external "Clear All Slicers" button, clicking the button on a different day restores the static date from when the bookmark was created, instead of recalculating based on the current date.

## Solution Overview

The solution implements **two-stage bookmark restore detection**:

1. **Early Detection** (lines 102-179): Detects stale bookmark dates before processing data bounds
2. **Fallback Detection** (lines 620-757): Catches cases where early detection didn't have data available

## Key Implementation Details

### 1. Bookmark Restore Detection

The solution detects bookmark restore by:
- Checking if external filters (`options.jsonFilters`) contain dates
- Comparing those dates with the **live preset calculation** (using `calculatePresetRange()`)
- If dates differ by more than 1 day, treating it as a stale bookmark restore
- Only processing if:
  - `recalculatePresetOnBookmark` setting is enabled (default: true)
  - Preset is time-based (today, yesterday, last7Days, etc.)
  - User has NOT manually selected dates (`hasManualSelection === false`)
  - Change is NOT user-initiated (`isUserInitiatedChange === false`)

### 2. Live Date Recalculation

When bookmark restore is detected:
- Recalculates preset range using `calculatePresetRange()` with current date
- Overrides `selectedMinDate` and `selectedMaxDate` with live dates
- Immediately applies filter using `applyDateFilter()`
- Updates React UI using `reactSliderWrapper.updateDates()`
- Sets `isRestoringBookmark = true` to prevent other logic from interfering

### 3. Manual Selection Protection

**Critical**: Manual calendar selections are ALWAYS respected:
- `hasManualSelection` flag prevents preset override
- Manual selections take precedence over bookmark restore
- Only programmatic changes (bookmark restore) can override presets when no manual selection exists

### 4. Flag Management

The solution uses several flags to prevent infinite loops and ensure correct behavior:

- `isRestoringBookmark`: Marks programmatic bookmark restore (not user action)
- `isUserInitiatedChange`: Marks user-initiated changes (calendar/slider)
- `hasManualSelection`: Tracks if user manually selected dates
- `shouldApplyFilter`: Controls when filters should be applied
- `detectedBookmarkRestore`: Tracks if bookmark restore was detected in this update cycle

### 5. VisualUpdateOptions.type Usage

**Note**: Power BI Visuals API doesn't provide a specific `VisualUpdateType` for bookmark restore. The solution detects it indirectly by:
- Analyzing `options.jsonFilters` for stale dates
- Comparing filter dates with live preset calculations
- Using timing and state flags to distinguish bookmark restore from other updates

## Code Flow

```
update(options) called
  ↓
Early Detection (lines 102-179)
  ├─ Check jsonFilters for stale dates
  ├─ Compare with live preset calculation
  └─ If stale: Override with live dates, apply filter, update UI
  ↓
Process data bounds and categories
  ↓
External Filter Handling (lines 620-757)
  ├─ Check if early detection missed it
  ├─ Re-check for stale dates with data bounds available
  └─ If stale: Override with live dates, apply filter, update UI
  ↓
Normal update flow continues
```

## Best Practices Implemented

1. **Separation of Concerns**: User-initiated vs programmatic changes are clearly separated
2. **Defensive Programming**: Multiple detection points ensure bookmark restore is caught
3. **Flag Safety**: Flags are cleared with timeouts to prevent infinite loops
4. **Manual Selection Priority**: User selections always take precedence
5. **React UI Sync**: React component is updated immediately when dates change
6. **Filter Application**: Filters are applied immediately to ensure other visuals update

## Testing Scenarios

1. ✅ **Day 1**: Create bookmark with "Today" preset → Works correctly
2. ✅ **Day 2**: Click Clear button → Should show Day 2's "Today", not Day 1's date
3. ✅ **Manual Selection**: User selects dates → Should never be overridden
4. ✅ **Clear After Manual**: Click Clear after manual selection → Should reset to preset (if no manual selection flag)
5. ✅ **Static Presets**: Non-time-based presets should work normally

## Configuration

The solution respects the `recalculatePresetOnBookmark` setting in the format pane:
- **Enabled** (default): Bookmark restore recalculates time-based presets
- **Disabled**: Bookmark restore uses stored dates (legacy behavior)

## Microsoft Visuals Best Practices

This implementation follows patterns used by Microsoft's built-in visuals:
- Store preset **intent** (preset name) rather than resolved dates
- Recalculate dynamic presets on bookmark restore
- Use flags to distinguish user actions from programmatic changes
- Apply filters immediately when dates change programmatically
- Sync UI components immediately after date changes



