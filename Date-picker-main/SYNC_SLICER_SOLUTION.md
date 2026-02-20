# Power BI Sync Slicer Solution

## Overview
This document explains the solution for fixing sync slicer issues in Power BI Custom Visual date range slicer with presets and manual date selection.

## Problem Statement

### SCENARIO 1: Page Sync (One Page Hidden)
- Page 1: Sync = ON, View = ON
- Page 2: Sync = ON, View = OFF
- **Issue**: Manual date selection doesn't persist when navigating back to Page 1
- **Root Cause**: Visual recomputes presets instead of honoring `options.jsonFilters`

### SCENARIO 2: Both Pages Visible
- Page 1: Sync = ON, View = ON
- Page 2: Sync = ON, View = ON
- **Issue**: Slicers overwrite each other (sync loops)
- **Root Cause**: No distinction between user-initiated changes and synced updates

## Key Concepts

### 1. Power BI Destroys Visuals on Page Navigation
**Why**: Power BI destroys and recreates visual instances when navigating between pages. Each page gets its own visual instance.

**Implication**: 
- Instance state is lost on navigation
- `options.jsonFilters` becomes the single source of truth for synced filters
- We must honor external filters from `jsonFilters` instead of recomputing

### 2. jsonFilters as Single Source of Truth
`options.jsonFilters` contains the current filter state across ALL pages. When Sync Slicers is enabled:
- User changes date on Page 1 → Filter appears in `jsonFilters`
- Page 2 receives this filter in its next `update()` call
- Page 2 must honor this filter, not recompute from preset

### 3. Update Types
The solution distinguishes between:
- **Initial Render**: First `update()` call after construction
- **User-Initiated Change**: User drags slider/selects dates → Apply filter
- **Sync Update**: External filter differs from what we last applied → Honor filter
- **Re-render/Navigation**: Visual recreated → Check jsonFilters first

## Solution Implementation

### State Tracking

```typescript
// New state variables
private lastAppliedFilterHash: string | null = null;  // Hash of filter we last applied
private isInitialRender: boolean = true;              // First update() call after construction
private isSyncUpdateFlag: boolean = false;            // Receiving external filter from sync
```

### Guard Conditions

#### 1. `isUserInitiatedChange`
- **True when**: User drags slider/selects dates in calendar
- **Action**: APPLY filter (don't honor external filter)
- **Set**: In `onDateChange` callback and `updateSelectedDates()`
- **Reset**: After 100-200ms timeout

#### 2. `isSyncUpdateFlag` 
- **True when**: 
  - External filter exists (not null)
  - We have a `lastAppliedFilterHash`
  - External filter hash differs from what we last applied
  - NOT a user-initiated change
- **Action**: HONOR external filter (don't apply new filter)
- **Set**: Via `checkIsSyncUpdate()` method

#### 3. `isInitialRender`
- **True when**: First `update()` call after construction
- **Action**: Apply preset if no external filter exists
- **Set**: Initially `true`, set to `false` after first `update()`

#### 4. `hasManualSelection`
- **True when**: User has manually selected dates (not from preset)
- **Action**: Prevent preset from overriding manual selection
- **Set**: In `onDateChange` callback and `updateSelectedDates()`

### Filter Handling Priority

The code follows this priority order in `update()` method:

1. **PRIORITY 1: Clear All detection** (highest)
   - If `clearAllPending` → Use live preset dates, ignore stale external filters

2. **PRIORITY 2: Sync update** 
   - If `externalFilterRange` exists AND `isSyncUpdateFlag` → Honor external filter
   - Set `hasManualSelection = true` to prevent preset override
   - Set `shouldApplyFilter = false` to avoid re-applying existing filter

3. **PRIORITY 3: External filter (not sync)**
   - If `externalFilterRange` exists but NOT sync update → Use preset if no manual selection, otherwise honor external filter

4. **PRIORITY 4: No external filter**
   - If no `externalFilterRange` → Compute from preset or use data bounds
   - On initial render with preset → Apply preset

### Preventing Sync Loops

Sync loops are prevented by:
1. **Tracking last applied filter hash**: When we apply a filter, we store its hash
2. **Comparing hashes**: If external filter hash matches our last applied hash → Not a sync update (it's our own filter propagating back)
3. **Not re-applying**: When `isSyncUpdateFlag = true`, we set `shouldApplyFilter = false` to avoid circular updates

### Code Flow

```typescript
public update(options: VisualUpdateOptions): void {
    // STEP 1: Detect update type early (before processing data)
    const incomingFilters = (options.jsonFilters as powerbi.IFilter[]) || [];
    const wasInitialRender = this.isInitialRender;
    if (this.isInitialRender) {
        this.isInitialRender = false;
    }

    // ... process data, calculate preset ranges ...

    // STEP 2: Handle external filters (jsonFilters as single source of truth)
    const externalFilterRange = this.getDateRangeFromFilters(incomingFilters, category.source);
    const externalFilterHash = this.createFilterHash(incomingFilters, category.source);
    this.isSyncUpdateFlag = this.checkIsSyncUpdate(externalFilterHash);

    // Apply priority-based filter handling (see above)
    
    // STEP 3: Update filter hash when applying filters
    if (this.shouldApplyFilter) {
        this.applyDateFilter(...);
        this.lastAppliedFilterHash = this.createFilterHash([appliedFilter], source);
    }
}
```

## Where Logic Belongs in `update(options)`

The sync handling logic belongs in this order within `update()`:

1. **Early (before data processing)**: 
   - Extract `incomingFilters` from `options.jsonFilters`
   - Track `isInitialRender` state

2. **After data bounds calculated**:
   - Calculate preset ranges (`presetRangeForClear`)
   - Extract external filter range from `incomingFilters`
   - Detect sync updates

3. **In external filter handling section** (around line 698-800):
   - Apply priority-based filter handling
   - Honor external filters when sync update detected
   - Apply presets only when no external filter exists

4. **In `applyDateFilter()` method**:
   - Update `lastAppliedFilterHash` after applying filter

5. **In `updateSelectedDates()` method**:
   - Update `lastAppliedFilterHash` after user-initiated changes

## Testing Scenarios

### Test 1: Page Navigation with Sync
1. Page 1: Select date range manually
2. Navigate to Page 2 (hidden)
3. Navigate back to Page 1
4. **Expected**: Manual selection persists (honored from jsonFilters)

### Test 2: Both Pages Visible
1. Page 1: Select date range
2. **Expected**: Page 2 updates to match (no sync loop)
3. Page 2: Select different date range
4. **Expected**: Page 1 updates to match (no sync loop)

### Test 3: Preset with Sync
1. Page 1: Set preset to "Today"
2. Navigate to Page 2
3. **Expected**: Page 2 shows "Today" range
4. Page 2: Manually select different dates
5. Navigate to Page 1
6. **Expected**: Page 1 shows manual selection from Page 2

## Key Methods

### `createFilterHash(filters, source)`
Creates a hash of filter state for comparison. Used to detect if external filter is the same as what we last applied.

### `checkIsSyncUpdate(externalFilterHash)`
Determines if an external filter represents a sync update from another page.

### `getDateRangeFromFilters(filters, source)`
Extracts date range from Power BI filter objects.

## Constraints Met

✅ No undocumented Power BI APIs (uses only `options.jsonFilters`, `applyJsonFilter()`)  
✅ Works with Sync Slicers + Page Navigation  
✅ Manual selection always persists  
✅ Prevents sync loops  


