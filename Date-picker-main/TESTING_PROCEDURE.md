# Testing Procedure for Bookmark Preset Fix

## Issue Fixed
When a bookmark with preset "yesterday" is created and attached to an external button, clicking it should recalculate "yesterday" based on the **current date**, not the date when the bookmark was created.

## Testing Steps

### Prerequisites
1. Build the visual: `npm run build` or `pbiviz package`
2. Import the visual into Power BI Desktop or Service
3. Have a report with date data

### Test Scenario 1: Basic "Yesterday" Preset Bookmark

**Day 1 (e.g., December 5, 2025):**
1. Open your Power BI report
2. In the date picker visual, select preset **"Yesterday"** from the format pane
3. Verify the date shows December 4, 2025 (yesterday's date)
4. Create a bookmark:
   - Go to View → Bookmarks → New bookmark
   - Name it "Yesterday Preset"
   - Make sure "Data" is selected in the bookmark options
5. Attach the bookmark to an external button:
   - Insert a button (or shape) in your report
   - Right-click the button → Action → Bookmark
   - Select "Yesterday Preset"

**Day 2 (e.g., December 6, 2025):**
1. Open the same Power BI report
2. Click the external button that triggers the bookmark
3. **Expected Result:** The date picker should show **December 5, 2025** (today's yesterday), NOT December 4, 2025 (the date from when bookmark was created)
4. Verify the filter is applied correctly to other visuals

### Test Scenario 2: Other Time-Based Presets

Test with these presets (create bookmarks for each):
- **"Today"** - Should always show current date
- **"Last 3 Days"** - Should show last 3 days from current date
- **"Last 7 Days"** - Should show last 7 days from current date
- **"Last 30 Days"** - Should show last 30 days from current date
- **"This Month"** - Should show from 1st of current month to today
- **"Last Month"** - Should show the previous month

**Steps:**
1. Create bookmark with each preset
2. Wait a day or change system date
3. Click bookmark button
4. Verify dates are recalculated based on current date

### Test Scenario 3: Clear Filter Button

1. Set preset to "Yesterday" in format pane
2. Create a bookmark with "Yesterday" preset
3. Manually change the date selection to something else
4. Click the bookmark button
5. **Expected:** Should restore to "Yesterday" (recalculated for current date)

### Test Scenario 4: Console Logging (For Debugging)

1. Open browser Developer Tools (F12) or Power BI Desktop Developer Tools
2. Go to Console tab
3. Click the bookmark button
4. Look for these log messages:
   - `"Restoring bookmark state:"` - Should show the bookmark state with `lastPreset`
   - `"Bookmark restore - preset detected:"` - Should show preset detection
   - `"Recalculating preset on bookmark restore:"` - Should show recalculation
   - `"clearSelectionWithCurrentData - preset source:"` - Should show which preset is being used
   - `"Preset recalculated for clear filter:"` - Should show the recalculated dates

### Expected Console Output Example

```
Restoring bookmark state: {selectedMinDate: "...", lastPreset: "yesterday", ...}
Bookmark restore - preset detected: {preset: "yesterday", isTimeBasedPreset: true, ...}
Recalculating preset on bookmark restore: {preset: "yesterday", currentDate: "2025-12-06T..."}
clearSelectionWithCurrentData - preset source: {usingBookmarkPreset: true, bookmarkPreset: "yesterday", ...}
Preset recalculated for clear filter: {preset: "yesterday", minDate: "2025-12-05T...", maxDate: "2025-12-05T..."}
```

### Troubleshooting

**If the bookmark still shows old dates:**
1. Check console logs to see if preset is being detected
2. Verify `lastPreset` is in the bookmark state
3. Check if `isTimeBasedPreset` is true
4. Verify `calculatePresetRange` is being called with current date

**If preset is not detected:**
1. Check that bookmark was created with a preset selected
2. Verify bookmark has "Data" option enabled
3. Check console for `"No preset in bookmark state"` message

**If dates are wrong:**
1. Check system date is correct
2. Verify data bounds are correct
3. Check if preset calculation logic is working (see `calculatePresetRange` method)

## Verification Checklist

- [ ] Bookmark with "Yesterday" preset recalculates correctly
- [ ] Bookmark with "Today" preset recalculates correctly  
- [ ] Bookmark with "Last 3 Days" preset recalculates correctly
- [ ] Bookmark with "Last 7 Days" preset recalculates correctly
- [ ] Bookmark with "Last 30 Days" preset recalculates correctly
- [ ] Bookmark with "This Month" preset recalculates correctly
- [ ] Bookmark with "Last Month" preset recalculates correctly
- [ ] Console logs show preset detection and recalculation
- [ ] Filter is applied correctly to other visuals after bookmark restore

## Notes

- Data-bound presets (minDate, maxDate) still respect the "Recalculate preset on bookmark restore" setting
- Time-based presets ALWAYS recalculate regardless of the setting
- The fix ensures presets are dynamic and use current date, not static dates from bookmark creation



