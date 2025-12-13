"use strict";

import powerbi from "powerbi-visuals-api";
import { FormattingSettingsService } from "powerbi-visuals-utils-formattingmodel";
import "./../style/visual.less";

import VisualConstructorOptions = powerbi.extensibility.visual.VisualConstructorOptions;
import VisualUpdateOptions = powerbi.extensibility.visual.VisualUpdateOptions;
import IVisual = powerbi.extensibility.visual.IVisual;
import DialogAction = powerbi.DialogAction;
import ISelectionManager = powerbi.extensibility.ISelectionManager;

import { VisualFormattingSettingsModel } from "./settings";
import { ReactSliderWrapper } from "./ReactWrapper";
import { DatePickerDialog, DatePickerDialogResult } from "./DatePickerDialog";

export class Visual implements IVisual {
    private target: HTMLElement;
    private formattingSettings: VisualFormattingSettingsModel;
    private formattingSettingsService: FormattingSettingsService;
    private host: powerbi.extensibility.visual.IVisualHost;  // 🔑 keep host reference
    private selectionManager: ISelectionManager;
    private boundContextMenuHandler?: (event: MouseEvent) => void;

    // 🔑 persist user-selected dates
    private dataMinDate: Date | null = null;  // Min date from data (slider bounds)
    private dataMaxDate: Date | null = null;  // Max date from data (slider bounds)
    private lastPresetSetting: string | null = null;
    private activePreset: string | null = null;
    private reactSliderWrapper: ReactSliderWrapper | null = null;

    // Track the actual selected dates for filtering (separate from slider state)
    private selectedMinDate: Date | null = null;
    private selectedMaxDate: Date | null = null;

    // Reference to current data source for universal filtering
    private currentDataSource: powerbi.DataViewMetadataColumn | null = null;

    // Flag to track when visual needs to re-render due to external changes
    private needsReRender: boolean = false;

    // Track data changes for Desktop compatibility
    private lastDataHash: string | null = null;

    // Flag to control when to apply filters (prevents interference with other slicers)
    private shouldApplyFilter: boolean = true;

    // Flag to track if the current change is user-initiated (slider/calendar change)
    private isUserInitiatedChange: boolean = false;

    // Flag to track if bookmark state is being restored (prevents override during restore)
    private isRestoringBookmark: boolean = false;

    // Flag to track if preset just changed (ensures filter is applied in Desktop)
    private presetJustChanged: boolean = false;

    // Flag to know when a Clear All slicers action was triggered externally
    private clearAllPending: boolean = false;

    // Flag to track if user has manually selected dates (prevents preset from overriding manual selection)
    private hasManualSelection: boolean = false;

    private presetRangeForClear: { from: Date; to: Date } | null = null;
    private defaultRangeForClear: { from: Date; to: Date } | null = null;


    constructor(options: VisualConstructorOptions) {
        this.target = options.element;
        this.host = options.host;   // ✅ store host
        this.formattingSettingsService = new FormattingSettingsService();
        this.formattingSettings = new VisualFormattingSettingsModel();
        this.selectionManager = this.host.createSelectionManager();
        // Right-click anywhere inside the visual to show the built-in context menu
        this.boundContextMenuHandler = (event: MouseEvent) => {
            event.preventDefault();
            this.selectionManager.showContextMenu(undefined, {
                x: event.clientX,
                y: event.clientY
            });
        };
        this.target.addEventListener("contextmenu", this.boundContextMenuHandler);

    }

    public update(options: VisualUpdateOptions): void {


        // Detect environment
        const isService = typeof window !== 'undefined' && window.location && window.location.hostname.includes('app.powerbi.com');
        const environment = isService ? "Power BI Service" : "Power BI Desktop";

        console.log("Visual update called with options:", {
            hasDataViews: !!(options.dataViews && options.dataViews[0]),
            dataViewCount: options.dataViews?.length || 0,
            jsonFilters: options.jsonFilters,
            type: options.type,
            environment: environment
        });

        if (options.dataViews && options.dataViews[0]) {
            this.formattingSettings =
                this.formattingSettingsService.populateFormattingSettingsModel(
                    VisualFormattingSettingsModel,
                    options.dataViews[0]
                );

            const selectedPreset = String(this.formattingSettings.presetsCard.preset.value);
            const selectedStyle = this.formattingSettings.presetsCard.selectionStyle.value;
            const dataView = options.dataViews[0];
            const category = dataView.categorical?.categories?.[0];
            let presetRangeForClear: { from: Date; to: Date } | null = null;
            let defaultRangeForClear: { from: Date; to: Date } | null = null;

            if (category && category.values && category.values.length > 0) {
                // If a Date Hierarchy level is used (Year/Quarter/Month/Day),
                // show a friendly message and stop – this visual needs the base Date column.
                if (category.source && this.isHierarchyField(category.source)) {
                    this.renderMessage("Please bind the base Date column, not the Date Hierarchy (Year/Quarter/Month/Day).From the Visual pane, right-click the Date field and select Date instead of Date Hierarchy.");
                    return;
                }

                // Always keep the latest source handy for preset-driven changes
                this.currentDataSource = category.source;

                console.log("Category data:", {
                    valuesCount: category.values.length,
                    firstFewValues: category.values.slice(0, 5),
                    source: category.source,
                    environment: environment
                });

                // For Desktop: Check if data has changed even if bounds look the same
                if (!isService && this.dataMinDate && this.dataMaxDate) {
                    const currentDataHash = category.values.join('|');
                    if (this.lastDataHash && this.lastDataHash !== currentDataHash) {
                        console.log("Desktop: Data content changed, forcing update");
                        this.needsReRender = true;
                    }
                    this.lastDataHash = currentDataHash;
                }

                // Parse values to Date[]
                const parsedDates: Date[] = category.values
                    .map(v => this.parseToDate(v))
                    .filter((d): d is Date => d !== null && !isNaN(d.getTime()));

                if (parsedDates.length > 0) {
                    let minDate = new Date(Math.min(...parsedDates.map(d => d.getTime())));
                    let maxDate = new Date(Math.max(...parsedDates.map(d => d.getTime())));

                    const dataBoundMin = new Date(minDate.getTime());
                    const dataBoundMax = new Date(maxDate.getTime());
                    presetRangeForClear = this.calculatePresetRange(selectedPreset);
                    defaultRangeForClear = { from: new Date(dataBoundMin), to: new Date(dataBoundMax) };

                    // Persist the latest preset/default ranges for use in clear/bookmark paths.
                    this.presetRangeForClear = presetRangeForClear;
                    this.defaultRangeForClear = defaultRangeForClear;

                    // If a clear-all or bookmark restore is underway, force the live preset range
                    // into the current selection so currentMin/Max reflect live dates (not bookmark dates).
                    const isTimeBasedPresetForClear =
                        selectedPreset &&
                        ["today", "yesterday", "last3days", "last7Days", "last30Days", "thisMonth", "lastMonth"].includes(selectedPreset);
                    if (this.clearAllPending && presetRangeForClear && isTimeBasedPresetForClear) {
                        this.selectedMinDate = presetRangeForClear.from;
                        this.selectedMaxDate = presetRangeForClear.to;
                        this.hasManualSelection = false;
                        this.needsReRender = true;
                    }

                    console.log("Data received:", {
                        dateCount: parsedDates.length,
                        minDate: minDate.toISOString(),
                        maxDate: maxDate.toISOString(),
                        currentDataMin: this.dataMinDate?.toISOString(),
                        currentDataMax: this.dataMaxDate?.toISOString()
                    });

                    // Strip times
                    minDate.setHours(0, 0, 0, 0);
                    maxDate.setHours(23, 59, 59, 999);

                    // Check if data bounds changed (external slicer filtered the data)
                    const boundsChanged = (!this.dataMinDate || !this.dataMaxDate) ||
                        (this.dataMinDate.getTime() !== dataBoundMin.getTime() || this.dataMaxDate.getTime() !== dataBoundMax.getTime());

                    console.log("Bounds change detection:", {
                        boundsChanged,
                        currentDataMin: this.dataMinDate?.toISOString(),
                        currentDataMax: this.dataMaxDate?.toISOString(),
                        newDataMin: minDate.toISOString(),
                        newDataMax: maxDate.toISOString(),
                        environment: environment
                    });

                    // Force bounds change detection in Desktop if data seems different
                    // Desktop sometimes doesn't trigger bounds change properly
                    if (!isService && !boundsChanged && this.dataMinDate && this.dataMaxDate) {
                        const timeDiff = Math.abs(this.dataMinDate.getTime() - minDate.getTime()) +
                            Math.abs(this.dataMaxDate.getTime() - maxDate.getTime());
                        if (timeDiff > 1000) { // More than 1 second difference
                            console.log("Desktop: Forcing bounds change due to time difference:", timeDiff);
                            // Force the bounds change
                            this.dataMinDate = minDate;
                            this.dataMaxDate = maxDate;
                            this.selectedMinDate = minDate;
                            this.selectedMaxDate = maxDate;
                            this.needsReRender = true;
                        }
                    }

                    // Handle external slicer changes (other slicers filtering data)
                    if (boundsChanged && !this.isRestoringBookmark) {
                        console.log("External slicer detected - bounds changed:", {
                            oldMin: this.dataMinDate?.toISOString(),
                            oldMax: this.dataMaxDate?.toISOString(),
                            newMin: minDate.toISOString(),
                            newMax: maxDate.toISOString()
                        });

                        // Check if this is initial load (dataMinDate is null)
                        const isInitialLoad = !this.dataMinDate || !this.dataMaxDate;

                        // Update data bounds to reflect external slicer changes
                        this.dataMinDate = minDate;
                        this.dataMaxDate = maxDate;

                        // If we have a current selection, clamp it to new bounds
                        if (this.selectedMinDate && this.selectedMaxDate) {
                            const clampedMin = new Date(Math.max(this.selectedMinDate.getTime(), minDate.getTime()));
                            const clampedMax = new Date(Math.min(this.selectedMaxDate.getTime(), maxDate.getTime()));
                            this.selectedMinDate = clampedMin;
                            this.selectedMaxDate = clampedMax;
                            console.log("Clamped selection to new bounds:", {
                                clampedMin: clampedMin.toISOString(),
                                clampedMax: clampedMax.toISOString()
                            });
                        } else {
                            // No current selection - check if preset should be applied (initial load)
                            if (isInitialLoad && presetRangeForClear && selectedPreset && selectedPreset !== "none") {
                                // Apply preset on initial load
                                this.selectedMinDate = presetRangeForClear.from;
                                this.selectedMaxDate = presetRangeForClear.to;
                                this.activePreset = selectedPreset;
                                this.lastPresetSetting = selectedPreset;
                                console.log("Initial load - applying preset:", {
                                    preset: selectedPreset,
                                    min: presetRangeForClear.from.toISOString(),
                                    max: presetRangeForClear.to.toISOString()
                                });
                                // Apply filter for initial preset load
                                if (category?.source) {
                                    this.isUserInitiatedChange = true;
                                    this.shouldApplyFilter = true;
                                    // Pass isPresetChange = true so it doesn't mark as manual selection
                                    this.updateSelectedDates(presetRangeForClear.from, presetRangeForClear.to, category.source, true);
                                    setTimeout(() => {
                                        this.isUserInitiatedChange = false;
                                    }, 100);
                                }
                            } else {
                                // No preset or not initial load - use data bounds
                                this.selectedMinDate = minDate;
                                this.selectedMaxDate = maxDate;
                                console.log("Set selection to filtered bounds:", {
                                    selectedMin: minDate.toISOString(),
                                    selectedMax: maxDate.toISOString()
                                });
                            }
                        }
                        this.needsReRender = true;
                        // Don't apply filter when external slicers change (unless it's initial load with preset)
                        // The filter should only be applied when user actively changes the date selection
                    } else if (boundsChanged && this.isRestoringBookmark) {
                        // During bookmark restore, just update data bounds but preserve restored selection
                        this.dataMinDate = minDate;
                        this.dataMaxDate = maxDate;
                        console.log("Bounds changed during bookmark restore - preserving restored selection");
                    }

                    // Handle preset changes (only when preset actually changes, and not during bookmark restore)
                    // Also handle initial load when preset is set but lastPresetSetting is null
                    const isPresetChange = this.lastPresetSetting !== selectedPreset;
                    const isInitialPresetLoad = !this.lastPresetSetting && selectedPreset && selectedPreset !== "none";

                    if ((isPresetChange || isInitialPresetLoad) && !this.isRestoringBookmark) {
                        console.log("Preset changed from", this.lastPresetSetting, "to", selectedPreset);
                        this.dataMinDate = dataBoundMin;
                        this.dataMaxDate = dataBoundMax;

                        // Determine the date range based on preset
                        let presetSelectionMin: Date;
                        let presetSelectionMax: Date;

                        if (presetRangeForClear && selectedPreset && selectedPreset !== "none") {
                            // Use preset range
                            presetSelectionMin = presetRangeForClear.from;
                            presetSelectionMax = presetRangeForClear.to;
                            this.activePreset = selectedPreset;
                        } else {
                            // "none" preset or no preset - use full data bounds
                            presetSelectionMin = dataBoundMin;
                            presetSelectionMax = dataBoundMax;
                            this.activePreset = null;
                        }

                        this.selectedMinDate = presetSelectionMin;
                        this.selectedMaxDate = presetSelectionMax;
                        this.lastPresetSetting = selectedPreset;

                        // Mark that preset just changed - this ensures filter is applied even in Desktop
                        this.presetJustChanged = true;

                        // Apply filter when preset changes (user-initiated action)
                        // This ensures the filter is applied immediately when preset changes in format pane
                        this.isUserInitiatedChange = true;
                        this.shouldApplyFilter = true;
                        this.needsReRender = true;

                        // Reset manual selection flag when preset changes from format pane
                        // This allows preset to override any previous manual selection
                        this.hasManualSelection = false;

                        if (category?.source) {
                            console.log("Preset change - applying filter via updateSelectedDates", {
                                preset: selectedPreset,
                                min: presetSelectionMin.toISOString(),
                                max: presetSelectionMax.toISOString()
                            });
                            // Always apply filter when preset changes from format pane
                            // Pass isPresetChange = true so it doesn't mark as manual selection
                            this.updateSelectedDates(presetSelectionMin, presetSelectionMax, category.source, true);

                            // Also apply filter directly here as backup for Desktop
                            // Desktop sometimes needs explicit filter application in the same update cycle
                            console.log("Preset change - applying filter directly as backup for Desktop");
                            this.applyDateFilter(category.source, presetSelectionMin, presetSelectionMax);
                        } else {
                            // Data source not available yet - will be applied in next update cycle
                            console.log("Preset changed but data source not available - will apply filter in next update");
                        }
                        // Reset the flag after a longer delay to allow filter application in Desktop
                        // Desktop update cycles can be slower, so we keep the flag longer
                        setTimeout(() => {
                            this.isUserInitiatedChange = false;
                            this.presetJustChanged = false;
                        }, 500);
                    } else if (!boundsChanged) {
                        // If no preset change and no bounds change, ensure we have valid dates but don't override user selection
                        if (!this.dataMinDate || !this.dataMaxDate) {
                            this.dataMinDate = dataBoundMin;
                            this.dataMaxDate = dataBoundMax;
                        }

                        // Check if preset is set and needs to be enforced
                        // This handles cases where preset is already set in format pane but selection doesn't match
                        // BUT only if user hasn't manually selected dates
                        if (presetRangeForClear && selectedPreset && selectedPreset !== "none" && !this.hasManualSelection) {
                            const presetMin = presetRangeForClear.from;
                            const presetMax = presetRangeForClear.to;

                            // Check if current selection doesn't match preset (or no selection exists)
                            const selectionMismatch = !this.selectedMinDate ||
                                !this.selectedMaxDate ||
                                this.selectedMinDate.getTime() !== presetMin.getTime() ||
                                this.selectedMaxDate.getTime() !== presetMax.getTime();

                            if (selectionMismatch) {
                                console.log("Preset is set but selection doesn't match - enforcing preset:", {
                                    preset: selectedPreset,
                                    currentMin: this.selectedMinDate?.toISOString(),
                                    currentMax: this.selectedMaxDate?.toISOString(),
                                    presetMin: presetMin.toISOString(),
                                    presetMax: presetMax.toISOString()
                                });

                                // Enforce preset range
                                this.selectedMinDate = presetMin;
                                this.selectedMaxDate = presetMax;
                                this.activePreset = selectedPreset;
                                this.lastPresetSetting = selectedPreset;

                                // Apply filter to enforce preset
                                this.presetJustChanged = true;
                                this.isUserInitiatedChange = true;
                                this.shouldApplyFilter = true;

                                if (category?.source) {
                                    console.log("Enforcing preset - applying filter");
                                    // Pass isPresetChange = true so it doesn't mark as manual selection
                                    this.updateSelectedDates(presetMin, presetMax, category.source, true);
                                    // Also apply directly as backup for Desktop
                                    this.applyDateFilter(category.source, presetMin, presetMax);
                                }

                                setTimeout(() => {
                                    this.isUserInitiatedChange = false;
                                    this.presetJustChanged = false;
                                }, 500);
                            } else {
                                // Selection matches preset - just ensure activePreset is set
                                this.activePreset = selectedPreset;
                                this.lastPresetSetting = selectedPreset;
                            }
                        } else if (this.hasManualSelection) {
                            // User has manually selected dates - don't enforce preset
                            console.log("Skipping preset enforcement - user has manually selected dates");
                        } else if (!this.selectedMinDate || !this.selectedMaxDate) {
                            // No preset and no selection - use data bounds
                            this.selectedMinDate = minDate;
                            this.selectedMaxDate = maxDate;
                        }
                    }

                    // General preset enforcement check - ensure preset is always enforced if set
                    // BUT only if user hasn't manually selected dates
                    // This runs after all other logic to catch any cases where preset wasn't applied
                    if (presetRangeForClear && selectedPreset && selectedPreset !== "none" && !this.isRestoringBookmark && !this.hasManualSelection) {
                        const presetMin = presetRangeForClear.from;
                        const presetMax = presetRangeForClear.to;

                        // Check if selection doesn't match preset
                        const needsEnforcement = !this.selectedMinDate ||
                            !this.selectedMaxDate ||
                            Math.abs(this.selectedMinDate.getTime() - presetMin.getTime()) > 1000 || // Allow 1 second tolerance
                            Math.abs(this.selectedMaxDate.getTime() - presetMax.getTime()) > 1000;

                        if (needsEnforcement && this.lastPresetSetting === selectedPreset) {
                            // Preset is set but selection doesn't match - enforce it
                            // Only if user hasn't manually selected dates
                            console.log("General preset enforcement - correcting selection to match preset:", {
                                preset: selectedPreset,
                                currentMin: this.selectedMinDate?.toISOString(),
                                currentMax: this.selectedMaxDate?.toISOString(),
                                presetMin: presetMin.toISOString(),
                                presetMax: presetMax.toISOString()
                            });

                            this.selectedMinDate = presetMin;
                            this.selectedMaxDate = presetMax;
                            this.activePreset = selectedPreset;

                            // Mark for filter application
                            this.presetJustChanged = true;
                            this.isUserInitiatedChange = true;
                            this.shouldApplyFilter = true;

                            // Apply filter directly as backup for Desktop
                            if (category?.source) {
                                console.log("General preset enforcement - applying filter directly");
                                this.applyDateFilter(category.source, presetMin, presetMax);
                            }
                        }
                    } else if (this.hasManualSelection) {
                        console.log("Skipping preset enforcement - user has manually selected dates");
                    }

                    // Force time-based presets to use live dates after bookmark restore.
                    // This prevents stale bookmark dates from persisting when the preset should be dynamic.
                    const isTimeBasedPresetForBookmark =
                        this.isRestoringBookmark &&
                        presetRangeForClear &&
                        selectedPreset &&
                        ["today", "yesterday", "last3days", "last7Days", "last30Days", "thisMonth", "lastMonth"].includes(selectedPreset);

                    if (isTimeBasedPresetForBookmark) {
                        console.log("Restoring bookmark with time-based preset - enforcing live preset range", {
                            preset: selectedPreset,
                            liveMin: presetRangeForClear.from.toISOString(),
                            liveMax: presetRangeForClear.to.toISOString()
                        });

                        this.selectedMinDate = presetRangeForClear.from;
                        this.selectedMaxDate = presetRangeForClear.to;
                        this.activePreset = selectedPreset;
                        this.lastPresetSetting = selectedPreset;
                        this.hasManualSelection = false;
                        this.shouldApplyFilter = true;
                        this.isUserInitiatedChange = true;
                        this.needsReRender = true;

                        if (category?.source) {
                            this.applyDateFilter(category.source, presetRangeForClear.from, presetRangeForClear.to);
                        }

                        // Clear bookmark restore flag once live dates are enforced so subsequent updates behave normally.
                        setTimeout(() => {
                            this.isRestoringBookmark = false;
                            this.isUserInitiatedChange = false;
                        }, 100);
                    }

                    // If we have a time-based preset and we are NOT in a user-initiated change,
                    // make sure selection matches the current (live) preset range. This prevents
                    // the selection from sticking to old bookmark dates.
                    const isTimeBasedPreset =
                        selectedPreset &&
                        ["today", "yesterday", "last3days", "last7Days", "last30Days", "thisMonth", "lastMonth"].includes(selectedPreset);
                    const presetRangeCurrent = presetRangeForClear;
                    const selectionIsStale =
                        isTimeBasedPreset &&
                        presetRangeCurrent &&
                        this.selectedMinDate &&
                        this.selectedMaxDate &&
                        (Math.abs(this.selectedMinDate.getTime() - presetRangeCurrent.from.getTime()) > 1000 ||
                            Math.abs(this.selectedMaxDate.getTime() - presetRangeCurrent.to.getTime()) > 1000);

                    // Only override when the change is not user-initiated AND either:
                    //  - user has not made a manual selection, OR
                    //  - we're in a bookmark/clear-all path (clearAllPending or isRestoringBookmark)
                    const canOverrideManual =
                        !this.hasManualSelection || this.clearAllPending || this.isRestoringBookmark;

                    if (isTimeBasedPreset && presetRangeCurrent && selectionIsStale && !this.isUserInitiatedChange && canOverrideManual) {
                        console.log("Time-based preset enforcement to live dates (non-user path)", {
                            preset: selectedPreset,
                            liveMin: presetRangeCurrent.from.toISOString(),
                            liveMax: presetRangeCurrent.to.toISOString(),
                            staleMin: this.selectedMinDate?.toISOString(),
                            staleMax: this.selectedMaxDate?.toISOString()
                        });

                        this.selectedMinDate = presetRangeCurrent.from;
                        this.selectedMaxDate = presetRangeCurrent.to;
                        this.activePreset = selectedPreset;
                        this.lastPresetSetting = selectedPreset;
                        this.hasManualSelection = false; // allow preset to own the state in this automated path
                        this.shouldApplyFilter = true;
                        this.needsReRender = true;

                        if (category?.source) {
                            this.applyDateFilter(category.source, presetRangeCurrent.from, presetRangeCurrent.to);
                        }
                    }

                    //  Apply filter to report (derive table/column from bound field)
                    if (this.dataMinDate && this.dataMaxDate) {
                        // Store the data source for universal filtering
                        this.currentDataSource = category.source;

                        const incomingFilters = (options.jsonFilters as powerbi.IFilter[]) || [];
                        const externalFilterRange = this.getDateRangeFromFilters(incomingFilters, category.source);
                        const filtersCleared = !externalFilterRange;

                        // Hard enforcement right before render: if coming from clear-all/bookmark and time-based preset,
                        // align selection and filter with live preset range even if currentMin/Max differ.
                        const isTimePresetNow =
                            selectedPreset &&
                            ["today", "yesterday", "last3days", "last7Days", "last30Days", "thisMonth", "lastMonth"].includes(selectedPreset);
                        const enforceLivePresetForClear =
                            this.clearAllPending &&
                            isTimePresetNow &&
                            !!presetRangeForClear &&
                            this.selectedMinDate &&
                            this.selectedMaxDate &&
                            (Math.abs(this.selectedMinDate.getTime() - presetRangeForClear.from.getTime()) > 1000 ||
                                Math.abs(this.selectedMaxDate.getTime() - presetRangeForClear.to.getTime()) > 1000);

                        if (enforceLivePresetForClear) {
                            console.log("Enforcing live preset after clear-all just before render", {
                                preset: selectedPreset,
                                liveMin: presetRangeForClear.from.toISOString(),
                                liveMax: presetRangeForClear.to.toISOString(),
                                currentMin: this.selectedMinDate.toISOString(),
                                currentMax: this.selectedMaxDate.toISOString()
                            });
                            this.selectedMinDate = presetRangeForClear.from;
                            this.selectedMaxDate = presetRangeForClear.to;
                            this.hasManualSelection = false;
                            this.shouldApplyFilter = true;
                            this.isUserInitiatedChange = true;
                            if (category?.source) {
                                this.applyDateFilter(category.source, presetRangeForClear.from, presetRangeForClear.to);
                            }
                            // Prevent repeated enforcement on next cycle
                            this.clearAllPending = false;
                            setTimeout(() => {
                                this.isUserInitiatedChange = false;
                            }, 50);
                        }

                        if (externalFilterRange && !this.isUserInitiatedChange && !this.isRestoringBookmark) {
                            // Don't override dates when restoring bookmark - let the bookmark restore logic handle it
                            const { minDate: externalMin, maxDate: externalMax } = externalFilterRange;
                            const selectionChanged =
                                !this.selectedMinDate ||
                                !this.selectedMaxDate ||
                                this.selectedMinDate.getTime() !== externalMin.getTime() ||
                                this.selectedMaxDate.getTime() !== externalMax.getTime();

                            if (selectionChanged) {
                                console.log("External filter detected via jsonFilters:", {
                                    externalMin: externalMin.toISOString(),
                                    externalMax: externalMax.toISOString()
                                });
                                this.selectedMinDate = externalMin;
                                this.selectedMaxDate = externalMax;
                                // Prevent immediately re-applying the same filter; wait for user action
                                this.shouldApplyFilter = false;
                                this.needsReRender = true;
                            }
                        } else if (externalFilterRange && this.isRestoringBookmark) {
                            // Power BI is trying to apply old filter dates from bookmark
                            // We need to ignore these and recalculate based on current date for time-based presets
                            console.log("Power BI attempting to apply old bookmark filter dates - will recalculate based on current date");
                            
                            // Check if we have a time-based preset that should be recalculated
                            const currentPreset = this.lastPresetSetting || String(this.formattingSettings.presetsCard.preset.value);
                            const timeBasedPresets = ["today", "yesterday", "last3days", "last7Days", "last30Days", "thisMonth", "lastMonth"];
                            const isTimeBased = currentPreset && timeBasedPresets.includes(currentPreset);
                            
                            if (isTimeBased) {
                                // Recalculate preset based on CURRENT date, ignore Power BI's old dates
                                const presetRange = this.calculatePresetRange(currentPreset);
                                if (presetRange) {
                                    console.log("Ignoring Power BI's old bookmark dates, using current date:", {
                                        oldMinDate: externalFilterRange.minDate.toISOString(),
                                        oldMaxDate: externalFilterRange.maxDate.toISOString(),
                                        newMinDate: presetRange.from.toISOString(),
                                        newMaxDate: presetRange.to.toISOString()
                                    });
                                    
                                    this.selectedMinDate = presetRange.from;
                                    this.selectedMaxDate = presetRange.to;
                                    this.presetRangeForClear = presetRange;
                                    this.hasManualSelection = false;
                                    
                                    // Apply the recalculated filter (not Power BI's old dates)
                                    if (category?.source) {
                                        this.shouldApplyFilter = true;
                                        this.isUserInitiatedChange = true;
                                        this.applyDateFilter(category.source, presetRange.from, presetRange.to);
                                        setTimeout(() => {
                                            this.isUserInitiatedChange = false;
                                        }, 100);
                                    }
                                    
                                    // Update React component with recalculated dates
                                    if (this.reactSliderWrapper) {
                                        this.reactSliderWrapper.updateDates(presetRange.from, presetRange.to, presetRange);
                                    }
                                }
                            } else {
                                // For non-time-based presets, still check if dates seem stale
                                // If preset exists, prefer recalculating it
                                if (currentPreset && currentPreset !== "none") {
                                    const presetRange = this.calculatePresetRange(currentPreset);
                                    if (presetRange) {
                                        console.log("Recalculating non-time-based preset instead of using Power BI's old dates");
                                        this.selectedMinDate = presetRange.from;
                                        this.selectedMaxDate = presetRange.to;
                                        this.presetRangeForClear = presetRange;
                                        
                                        if (category?.source) {
                                            this.shouldApplyFilter = true;
                                            this.isUserInitiatedChange = true;
                                            this.applyDateFilter(category.source, presetRange.from, presetRange.to);
                                            setTimeout(() => {
                                                this.isUserInitiatedChange = false;
                                            }, 100);
                                        }
                                        
                                        if (this.reactSliderWrapper) {
                                            this.reactSliderWrapper.updateDates(presetRange.from, presetRange.to, presetRange);
                                        }
                                    }
                                }
                            }
                        } else if (!this.isUserInitiatedChange && filtersCleared && this.dataMinDate && this.dataMaxDate) {
                            // External "Clear all slicers" detected
                            if (!this.clearAllPending) {
                                console.log("Clear all slicers detected - resetting to preset or data bounds");
                                this.clearAllPending = true;

                                // Get the current preset
                                const currentPreset = this.lastPresetSetting || String(this.formattingSettings.presetsCard.preset.value);

                                // Recalculate preset range based on CURRENT date (for time-based presets)
                                const presetRange = (currentPreset && currentPreset !== "none")
                                    ? this.calculatePresetRange(currentPreset)
                                    : null;

                                if (presetRange) {
                                    // Use recalculated preset range (LIVE dates for today/yesterday/etc)
                                    this.selectedMinDate = presetRange.from;
                                    this.selectedMaxDate = presetRange.to;
                                    this.activePreset = currentPreset;
                                    this.lastPresetSetting = currentPreset;
                                    console.log("External clear - reset to preset:", currentPreset, {
                                        liveMin: presetRange.from.toISOString(),
                                        liveMax: presetRange.to.toISOString()
                                    });
                                } else {
                                    // No preset - use full data bounds
                                    this.selectedMinDate = this.dataMinDate;
                                    this.selectedMaxDate = this.dataMaxDate;
                                    this.activePreset = null;
                                    console.log("External clear - reset to data bounds");
                                }

                                // Clear manual selection flag
                                this.hasManualSelection = false;

                                // Apply filter to other visuals
                                if (category?.source) {
                                    this.shouldApplyFilter = true;
                                    this.isUserInitiatedChange = true;
                                    this.applyDateFilter(category.source, this.selectedMinDate, this.selectedMaxDate);
                                    setTimeout(() => {
                                        this.isUserInitiatedChange = false;
                                    }, 100);
                                }

                                // Mark for re-render
                                this.needsReRender = true;

                                // Immediately push updated dates into React props so currentMin/Max reflect the live preset
                                if (this.reactSliderWrapper && this.selectedMinDate && this.selectedMaxDate) {
                                    this.reactSliderWrapper.updateDates(
                                        this.selectedMinDate,
                                        this.selectedMaxDate,
                                        presetRangeForClear
                                    );
                                }

                                // Clear the clearAllPending flag after processing
                                setTimeout(() => {
                                    this.clearAllPending = false;
                                }, 200);
                            }
                            // CRITICAL: Don't return here - let code continue to renderDateCard()
                        } else {
                            this.clearAllPending = false;
                        }


                        // Apply filter if:
                        // 1. Preset just changed (highest priority - always apply filter for preset changes in Desktop/Service)
                        //    This bypasses all other checks to ensure preset changes always filter
                        // 2. Bounds haven't changed (not an external slicer change), OR
                        // 3. This is a user-initiated change (user changed slider/calendar), OR
                        // 4. Initial load with selected dates, OR
                        // 5. Bookmark is being restored (need to apply filter when data source becomes available)
                        // BUT NOT when filters were just cleared externally AND it's not user-initiated AND not preset change
                        // AND only if shouldApplyFilter is true (explicitly set based on context)
                        const shouldApply = this.shouldApplyFilter &&
                            this.selectedMinDate &&
                            this.selectedMaxDate &&
                            (this.presetJustChanged || // Preset change always applies filter - bypasses all other checks
                                (!boundsChanged || this.isUserInitiatedChange || this.isRestoringBookmark) &&
                                (this.isUserInitiatedChange || !filtersCleared || this.isRestoringBookmark));

                        // For preset changes, force filter application even if other conditions might prevent it
                        // This is critical for Desktop where timing can be different
                        const forceApplyForPreset = this.presetJustChanged && this.selectedMinDate && this.selectedMaxDate;

                        // CRITICAL: If restoring bookmark and dates are set, ALWAYS apply filter
                        // This ensures filter is applied even if other conditions might prevent it
                        const forceApplyForBookmark = this.isRestoringBookmark &&
                            this.selectedMinDate &&
                            this.selectedMaxDate &&
                            this.shouldApplyFilter;

                        // Apply filter if conditions are met OR if preset just changed (force apply) OR if bookmark restore
                        if (shouldApply || forceApplyForPreset || forceApplyForBookmark) {
                            const filterMinDate = this.selectedMinDate;
                            const filterMaxDate = this.selectedMaxDate;
                            // Ensure filter is enabled when applying from update method
                            this.shouldApplyFilter = true;
                            console.log("Applying filter from update method:", {
                                isUserInitiated: this.isUserInitiatedChange,
                                isRestoringBookmark: this.isRestoringBookmark,
                                presetJustChanged: this.presetJustChanged,
                                forceApplyForPreset: forceApplyForPreset,
                                forceApplyForBookmark: forceApplyForBookmark,
                                boundsChanged,
                                filterMinDate: filterMinDate.toISOString(),
                                filterMaxDate: filterMaxDate.toISOString()
                            });
                            this.applyDateFilter(category.source, filterMinDate, filterMaxDate);

                            // If this was a bookmark restore and filter is now applied, we can clear the flag
                            if (this.isRestoringBookmark) {
                                // Filter applied successfully, clear the restore flag after a short delay
                                setTimeout(() => {
                                    this.isRestoringBookmark = false;
                                    this.isUserInitiatedChange = false;
                                    console.log("Bookmark restore completed - flags cleared");
                                }, 100);
                            }

                            // If preset just changed and filter is now applied, clear the flag after a delay
                            // Keep it a bit longer to ensure Desktop update cycles can use it
                            if (this.presetJustChanged) {
                                console.log("Preset change filter applied successfully in update cycle");
                                // Don't clear immediately - let the timeout handle it to ensure Desktop gets it
                            }
                        } else {
                            console.log("Skipping filter in update method:", {
                                isUserInitiated: this.isUserInitiatedChange,
                                isRestoringBookmark: this.isRestoringBookmark,
                                presetJustChanged: this.presetJustChanged,
                                boundsChanged,
                                hasSelectedDates: !!(this.selectedMinDate && this.selectedMaxDate),
                                shouldApplyFilter: this.shouldApplyFilter
                            });
                        }

                        // Display
                        const formatDate = (date: Date) =>
                            date.toLocaleDateString("en-GB", {
                                day: "2-digit",
                                month: "short",
                                year: "numeric"
                            });

                        this.renderDateCard(
                            this.dataMinDate,
                            this.dataMaxDate,
                            formatDate(this.dataMinDate),
                            formatDate(this.dataMaxDate),
                            category.source,
                            presetRangeForClear,
                            defaultRangeForClear
                        );

                        // Handle re-rendering after external filter changes
                        if (this.needsReRender && this.reactSliderWrapper) {
                            console.log("Re-rendering slider with new dates:", {
                                selectedMin: this.selectedMinDate?.toISOString(),
                                selectedMax: this.selectedMaxDate?.toISOString(),
                                dataMin: this.dataMinDate?.toISOString(),
                                dataMax: this.dataMaxDate?.toISOString()
                            });
                            this.reactSliderWrapper.updateDates(this.selectedMinDate || this.dataMinDate, this.selectedMaxDate || this.dataMaxDate);
                            this.needsReRender = false;
                        }
                    }
                } else {
                    this.renderMessage("No valid dates found");
                }
            } else {
                this.renderMessage("Please select a Date field");
            }
        } else {
            this.renderMessage("No data available");
        }
    }

    // Helper method to calculate preset date range
    private calculatePresetRange(preset: string): { from: Date; to: Date } | null {
        if (!preset || preset === "none") {
            return null;
        }

        const now = new Date();
        let presetMin: Date;
        let presetMax: Date;

        switch (preset) {
            case "today":
                presetMin = new Date(now);
                presetMax = new Date(now);
                break;
            case "yesterday":
                presetMin = new Date();
                presetMin.setDate(now.getDate() - 1);
                presetMax = new Date();
                presetMax.setDate(now.getDate() - 1);
                break;
            case "last3days":
                presetMin = new Date();
                presetMin.setDate(now.getDate() - 2);
                presetMax = new Date();
                break;
            case "thisMonth":
                presetMin = new Date(now.getFullYear(), now.getMonth(), 1);
                presetMax = new Date();
                break;
            case "lastMonth":
                presetMin = new Date(now.getFullYear(), now.getMonth() - 1, 1);
                presetMax = new Date(now.getFullYear(), now.getMonth(), 0);
                break;
            case "last7Days":
                presetMin = new Date();
                presetMin.setDate(now.getDate() - 7);
                presetMax = new Date();
                break;
            case "last30Days":
                presetMin = new Date();
                presetMin.setDate(now.getDate() - 30);
                presetMax = new Date();
                break;
            case "minDate":
                presetMin = this.dataMinDate ? new Date(this.dataMinDate) : new Date(now);
                presetMax = new Date(presetMin);
                break;
            case "maxDate":
                presetMax = this.dataMaxDate ? new Date(this.dataMaxDate) : new Date(now);
                presetMin = new Date(presetMax);
                break;
            default:
                return null;
        }

        // Normalize times
        presetMin.setHours(0, 0, 0, 0);
        presetMax.setHours(23, 59, 59, 999);

        // Clamp preset range to current data bounds if available
        // This ensures that for live data, presets respect the latest data range
        if (this.dataMinDate && this.dataMaxDate) {
            // Clamp min date to be within data bounds
            if (presetMin.getTime() < this.dataMinDate.getTime()) {
                presetMin = new Date(this.dataMinDate);
                presetMin.setHours(0, 0, 0, 0);
            }
            // Clamp max date to be within data bounds
            if (presetMax.getTime() > this.dataMaxDate.getTime()) {
                presetMax = new Date(this.dataMaxDate);
                presetMax.setHours(23, 59, 59, 999);
            }
            // Ensure min is not greater than max
            if (presetMin.getTime() > presetMax.getTime()) {
                presetMin = new Date(presetMax);
                presetMin.setHours(0, 0, 0, 0);
            }
        }

        return { from: presetMin, to: presetMax };
    }

    private getDateRangeFromFilters(
        filters: powerbi.IFilter[],
        source?: powerbi.DataViewMetadataColumn
    ): { minDate: Date; maxDate: Date } | null {
        if (!filters || !filters.length || !source) {
            return null;
        }

        const { tableName, columnName } = this.extractTableAndColumn(source);

        for (const filter of filters) {
            const advancedFilter = filter as any;
            const target = advancedFilter?.target;
            if (!target || target.table !== tableName || target.column !== columnName) {
                continue;
            }

            const conditions = advancedFilter?.conditions;
            if (!Array.isArray(conditions) || conditions.length === 0) {
                continue;
            }

            let minDate: Date | null = null;
            let maxDate: Date | null = null;

            for (const condition of conditions) {
                if (!condition || !condition.operator) continue;
                const operator = String(condition.operator).toLowerCase();
                const value = condition.value;
                const parsedDate = this.parseToDate(value);
                if (!parsedDate) continue;

                if (operator.includes("greater")) {
                    minDate = parsedDate;
                } else if (operator.includes("less")) {
                    maxDate = parsedDate;
                }
            }

            if (minDate && maxDate) {
                return { minDate, maxDate };
            }
        }

        return null;
    }

    private normalizeDateRange(minDate: Date, maxDate: Date): { normalizedMin: Date; normalizedMax: Date } {
        const normalizedMin = new Date(minDate);
        normalizedMin.setHours(0, 0, 0, 0);

        const normalizedMax = new Date(maxDate);
        normalizedMax.setHours(23, 59, 59, 999);

        if (normalizedMax.getTime() < normalizedMin.getTime()) {
            normalizedMax.setTime(normalizedMin.getTime());
        }

        return { normalizedMin, normalizedMax };
    }

    // 🔑 Centralized method to update dates and apply filter
    private updateSelectedDates(minDate: Date, maxDate: Date, source?: powerbi.DataViewMetadataColumn, isPresetChange: boolean = false): void {
        const { normalizedMin, normalizedMax } = this.normalizeDateRange(minDate, maxDate);

        console.log("updateSelectedDates called - User initiated change:", {
            minDate: normalizedMin.toISOString(),
            maxDate: normalizedMax.toISOString(),
            source: source?.queryName,
            isPresetChange: isPresetChange
        });

        this.selectedMinDate = normalizedMin;
        this.selectedMaxDate = normalizedMax;

        // Mark this as a user-initiated change - filter MUST be applied
        this.isUserInitiatedChange = true;
        this.shouldApplyFilter = true;

        // If this is a manual selection (not from preset change), mark it
        if (!isPresetChange) {
            this.hasManualSelection = true;
            console.log("Manual selection detected - preset will not override this");
        }

        // Use provided source or current data source
        const dataSource = source || this.currentDataSource;
        if (dataSource) {
            // Always apply filter for user-initiated changes
            this.applyDateFilter(dataSource, normalizedMin, normalizedMax);
        } else {
            console.warn("No data source available when trying to apply filter");
        }

        // Date display is now handled by React Day Picker component

        // Update slider if it exists
        if (this.reactSliderWrapper) {
            this.reactSliderWrapper.updateDates(normalizedMin, normalizedMax);
        }

        // Reset the flag after a short delay to allow update cycle to complete
        setTimeout(() => {
            this.isUserInitiatedChange = false;
        }, 100);
    }

    // 🔑 Apply JSON filter to whole report
    private applyDateFilter(source: powerbi.DataViewMetadataColumn, minDate: Date, maxDate: Date): void {
        const { tableName, columnName } = this.extractTableAndColumn(source);

        // Ensure dates are properly formatted
        const minDateStr = minDate.toISOString();
        const maxDateStr = maxDate.toISOString();

        // Use advanced filter for date ranges
        const advancedFilter: powerbi.IFilter = {
            $schema: "http://powerbi.com/product/schema#advanced",
            target: { table: tableName, column: columnName },
            logicalOperator: "And",
            conditions: [
                { operator: "GreaterThanOrEqual", value: minDateStr },
                { operator: "LessThanOrEqual", value: maxDateStr }
            ]
        };

        // Apply filter only when user actively changes the date slicer
        // This prevents automatic filter application from interfering with other slicers
        if (this.shouldApplyFilter) {
            console.log("Applying date filter to other visuals:", {
                tableName,
                columnName,
                minDate: minDateStr,
                maxDate: maxDateStr
            });
            // Use merge to combine with other slicers' filters
            // This ensures the filter is applied to other visuals while respecting other filters
            this.host.applyJsonFilter(
                advancedFilter,
                "general",
                "filter",
                powerbi.FilterAction.merge
            );
        } else {
            console.log("Filter application skipped - shouldApplyFilter is false");
        }
    }

    // Derive table and column reliably from metadata
    private extractTableAndColumn(source: powerbi.DataViewMetadataColumn): { tableName: string; columnName: string } {
        const queryName = source?.queryName || "";

        // Match formats like Table.Column
        let match = /^(?<table>[^\.\[]+)\.(?<column>.+)$/.exec(queryName);
        if (match?.groups) {
            return { tableName: match.groups.table, columnName: match.groups.column };
        }

        // Match formats like Table[Column]
        match = /^(?<table>[^\[]+)\[(?<column>[^\]]+)\]$/.exec(queryName);
        if (match?.groups) {
            return { tableName: match.groups.table, columnName: match.groups.column };
        }

        // Fallbacks: try explicit props if available
        const tableFallback = (source as any)?.table || "";
        const columnFallback = source?.displayName || queryName || "Date";
        return { tableName: tableFallback, columnName: columnFallback };
    }

    // Detect if the bound field is a Date Hierarchy level instead of the base Date column.
    // We try multiple heuristics:
    //  - queryName patterns like: Table[Date].[Year], Table[Date].[Month], etc.
    //  - display names like "Year", "Quarter", "Month", "Day" where the underlying type is not a date
    private isHierarchyField(source: powerbi.DataViewMetadataColumn): boolean {
        const q = source?.queryName || "";
        const displayName = source?.displayName || "";
        const type: any = (source as any)?.type || {};
        const isDateType = !!(type.dateTime || type.date);

        // Pattern 1: queryName contains hierarchy separator "].["
        const hasHierarchySeparator = q.indexOf("].[") !== -1;

        // Pattern 2: queryName explicitly ends with a known level name
        const hasLevelInQuery =
            /\.\[?(Year|Quarter|Month|Day)\]?$/i.test(q) ||
            /\]\.\[(Year|Quarter|Month|Day)\]$/i.test(q);

        // Pattern 3: display name is a typical level name and type is not a date
        const isLevelDisplayName = /^(year|quarter|month|day)$/i.test(displayName);
        const levelNameNonDate = isLevelDisplayName && !isDateType;

        const isHierarchy = !!(hasHierarchySeparator || hasLevelInQuery || levelNameNonDate);

        console.log("Hierarchy detection:", {
            queryName: q,
            displayName,
            type,
            isDateType,
            hasHierarchySeparator,
            hasLevelInQuery,
            isLevelDisplayName,
            levelNameNonDate,
            isHierarchy
        });

        return isHierarchy;
    }

    // 🔑 Robust parser supporting multiple formats
    private parseToDate(value: any): Date | null {
        if (value instanceof Date && !isNaN(value.getTime())) return value;

        // Numbers: epoch ms or Excel serial days
        if (typeof value === "number") {
            // Heuristic: small numbers likely Excel serial (days since 1899-12-30)
            if (value < 10_000_000_000) {
                const excelEpoch = new Date(Date.UTC(1899, 11, 30));
                const date = new Date(excelEpoch.getTime() + value * 24 * 60 * 60 * 1000);
                return isNaN(date.getTime()) ? null : date;
            }
            const d = new Date(value);
            return isNaN(d.getTime()) ? null : d;
        }

        if (typeof value === "string") {
            const trimmed = value.trim();

            // YYYY-MM-DD or YYYY/MM/DD
            let m = /^(\d{4})[-\/](\d{2})[-\/](\d{2})$/.exec(trimmed);
            if (m) {
                const d = new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]));
                return isNaN(d.getTime()) ? null : d;
            }

            // DD/MM/YYYY
            m = /^(\d{2})\/(\d{2})\/(\d{4})$/.exec(trimmed);
            if (m) {
                const d = new Date(Number(m[3]), Number(m[2]) - 1, Number(m[1]));
                return isNaN(d.getTime()) ? null : d;
            }

            // Try native parser (handles ISO and many locales)
            const parsed = new Date(trimmed);
            return isNaN(parsed.getTime()) ? null : parsed;
        }

        return null;
    }



    private renderDateCard(
        minDate: Date,
        maxDate: Date,
        minDateStr: string,
        maxDateStr: string,
        source: powerbi.DataViewMetadataColumn,
        presetRangeForClear?: { from: Date; to: Date } | null,
        defaultRangeForClear?: { from: Date; to: Date } | null
    ): void {

        this.target.innerHTML = "";
        const card = document.createElement("div");
        card.style.border = "1px solid #ccc";
        card.style.borderRadius = "8px";
        card.style.padding = "4px 0";
        card.style.textAlign = "center";
        card.style.background = (this.formattingSettings.dataPointCard.cardColor?.value?.value as string) || "#ffffff";
        card.style.fontFamily = "Segoe UI, sans-serif";

        if (this.formattingSettings.presetsCard.toShowHeader.value) {
            const title = document.createElement("div");
            title.textContent = "Mahaaya Super Date Slicer";
            title.style.fontWeight = "bold";
            title.style.marginBottom = "5px";
            card.appendChild(title);
        }

        const popupOnly = !!this.formattingSettings.presetsCard.toggleOption.value;
        if (popupOnly) {
            const calendarButton = document.createElement("button");
            calendarButton.textContent = "Open Calendar";
            calendarButton.style.marginTop = "10px";
            calendarButton.style.padding = "10px 14px";
            calendarButton.style.borderRadius = "10px";
            calendarButton.style.border = "1px solid #ccc";
            calendarButton.style.background = "#fff";
            calendarButton.style.cursor = "pointer";
            calendarButton.onclick = () => {
                this.openDatePickerDialog(
                    this.selectedMinDate || minDate,
                    this.selectedMaxDate || maxDate,
                    minDate,
                    maxDate
                );
            };
            card.appendChild(calendarButton);
        } else {
            const rangeSliderContainer = document.createElement("div");
            rangeSliderContainer.className = "range-slider-container";

            if (this.reactSliderWrapper) {
                this.reactSliderWrapper.destroy();
            }

            this.reactSliderWrapper = new ReactSliderWrapper(rangeSliderContainer);

            const inputFontSize = Number(this.formattingSettings.dataPointCard.fontSize.value) || 18;
            const inputFontColor = (this.formattingSettings.dataPointCard.fontColor?.value?.value as string) || "#000000";
            const inputBoxColor = (this.formattingSettings.dataPointCard.dateBoxColor?.value?.value as string) || "#ffffff";

            const selectedPreset = String(this.formattingSettings.presetsCard.preset.value);
            const isTimeBasedPreset = selectedPreset && selectedPreset !== "none" &&
                ["today", "yesterday", "last3days", "last7Days", "last30Days", "thisMonth", "lastMonth"].includes(selectedPreset);

            let currentMinToUse = this.selectedMinDate || minDate;
            let currentMaxToUse = this.selectedMaxDate || maxDate;

            console.log("currentMaxToUse", currentMaxToUse);
            console.log("currentMinToUse", currentMinToUse);

            // --- Prefer live preset range during clear/bookmark restore paths ---
            // Use preset range when coming from clear-all/bookmark restore even if isRestoringBookmark already cleared.
            const isClearOrBookmark = this.isRestoringBookmark || this.clearAllPending;

            // Also check if presetRangeForClear differs from selectedMinDate/selectedMaxDate (external clear case)
            // This ensures that when clearSelection() is called, we use the correct preset range even if selectedMinDate/selectedMaxDate
            // still have old bookmark dates
            const presetDiffersFromSelection = presetRangeForClear &&
                this.selectedMinDate && this.selectedMaxDate &&
                (Math.abs(this.selectedMinDate.getTime() - presetRangeForClear.from.getTime()) > 1000 ||
                    Math.abs(this.selectedMaxDate.getTime() - presetRangeForClear.to.getTime()) > 1000);

            // Apply preset range if: (1) clear/bookmark path, OR (2) preset differs from selection (external clear)
            // For time-based presets, always prefer preset range during clear/bookmark
            // For other presets, also check if preset differs from selection
            const shouldUsePresetRange = presetRangeForClear &&
                !this.hasManualSelection &&
                ((isTimeBasedPreset && (isClearOrBookmark || presetDiffersFromSelection)) ||
                    (!isTimeBasedPreset && presetDiffersFromSelection && selectedPreset && selectedPreset !== "none"));

            if (shouldUsePresetRange) {
                currentMinToUse = presetRangeForClear.from;
                currentMaxToUse = presetRangeForClear.to;
                this.selectedMinDate = currentMinToUse;
                this.selectedMaxDate = currentMaxToUse;

                // Also push the live dates into the display props so currentMin/Max reflect live, not bookmark dates.
                minDate = presetRangeForClear.from < minDate ? presetRangeForClear.from : minDate;
                maxDate = presetRangeForClear.to > maxDate ? presetRangeForClear.to : maxDate;

                if (this.currentDataSource) {
                    this.applyDateFilter(this.currentDataSource, currentMinToUse, currentMaxToUse);
                }

                // Once applied, clear flags so subsequent updates behave normally
                this.isRestoringBookmark = false;
                this.clearAllPending = false;
                this.hasManualSelection = false;
            }

            // Final safeguard: after external clear, if the live preset range differs from the current
            // min/max being shown, force render/filter to the preset range (not the stale currentMin/Max).
            const shouldForcePresetAfterClear =
                this.clearAllPending &&
                presetRangeForClear &&
                isTimeBasedPreset &&
                (Math.abs(currentMinToUse.getTime() - presetRangeForClear.from.getTime()) > 1000 ||
                    Math.abs(currentMaxToUse.getTime() - presetRangeForClear.to.getTime()) > 1000);

            if (shouldForcePresetAfterClear) {
                currentMinToUse = presetRangeForClear.from;
                currentMaxToUse = presetRangeForClear.to;
                this.selectedMinDate = currentMinToUse;
                this.selectedMaxDate = currentMaxToUse;
                this.hasManualSelection = false;

                // Ensure filter aligns with preset range
                if (this.currentDataSource) {
                    this.applyDateFilter(this.currentDataSource, currentMinToUse, currentMaxToUse);
                }
            }

            // Slider / inline calendar
            this.reactSliderWrapper.render({
                minDate,
                maxDate,
                currentMinDate: currentMinToUse,
                currentMaxDate: currentMaxToUse,
                onDateChange: (newMinDate: Date, newMaxDate: Date) => {
                    // Mark as manual selection
                    this.hasManualSelection = true;
                    this.selectedMinDate = newMinDate;
                    this.selectedMaxDate = newMaxDate;

                    if (this.currentDataSource) {
                        this.applyDateFilter(this.currentDataSource, newMinDate, newMaxDate);
                    }
                },
                formatDateForSlider: this.formatDateForSlider.bind(this),
                onOpenDialog: this.openDatePickerDialog.bind(this),
                datePickerType: this.formattingSettings.presetsCard.selectionStyle.value,
                inputFontSize,
                inputFontColor,
                inputBoxColor,
                presetRange: presetRangeForClear || null,
                defaultRange: defaultRangeForClear || null
            });

            card.appendChild(rangeSliderContainer);
        }

        this.target.appendChild(card);
    }




    private formatDateForSlider(date: Date): string {
        return date.toLocaleDateString("en-US", {
            month: "short",
            day: "numeric",
            year: "2-digit"
        });
    }


    // Method to open Power BI dialog for date selection
    public openDatePickerDialog(fromDate: Date, toDate: Date, minDate: Date, maxDate: Date): void {
        // Check if dialog is allowed in current environment
        if (!this.host.hostCapabilities.allowModalDialog) {
            console.warn("Modal dialogs are not allowed in this environment");
            return;
        }

        const dialogActionsButtons = [DialogAction.OK, DialogAction.Cancel];

        const initialDialogState = {
            fromDate: fromDate,
            toDate: toDate,
            minDate: minDate,
            maxDate: maxDate
        };

        const position = {
            type: 0 // Center position
        };

        // Dialog size - width fixed at 800px, height at 92vh
        const viewportHeight = typeof window !== 'undefined' ? window.innerHeight : 1080;
        const dialogHeight = viewportHeight * 0.92;

        const size = { width: 800, height: 460 };
        const dialogOptions = {
            actionButtons: dialogActionsButtons,
            size: size,
            position: position,
            title: ""
        };

        this.host.openModalDialog(DatePickerDialog.id, dialogOptions, initialDialogState)
            .then(ret => this.handleDialogResult(ret))
            .catch(error => this.handleDialogError(error));
    }

    // Handle dialog result
    private handleDialogResult(result: powerbi.extensibility.visual.ModalDialogResult): void {
        if (result.actionId === DialogAction.OK || result.actionId === DialogAction.Close) {
            const resultState = result.resultState as DatePickerDialogResult;
            if (resultState && (resultState.fromDate || resultState.toDate)) {
                const fallbackDate = resultState.fromDate || resultState.toDate;
                const selectedFromDate = resultState.fromDate
                    ? new Date(resultState.fromDate)
                    : fallbackDate
                        ? new Date(fallbackDate)
                        : null;
                const selectedToDate = resultState.toDate
                    ? new Date(resultState.toDate)
                    : fallbackDate
                        ? new Date(fallbackDate)
                        : null;

                if (selectedFromDate && selectedToDate) {
                    // Update selected dates and apply filter
                    this.updateSelectedDates(selectedFromDate, selectedToDate);
                }
            }
        }
    }

    // Handle dialog error
    private handleDialogError(error: any): void {
        console.error("Dialog error:", error);
    }

    private renderMessage(msg: string): void {
        this.target.innerHTML = "";
        const p = document.createElement("p");
        p.textContent = msg;
        p.style.color = "red";
        p.style.fontFamily = "Segoe UI, sans-serif";
        this.target.appendChild(p);
    }

    public getFormattingModel(): powerbi.visuals.FormattingModel {
        return this.formattingSettingsService.buildFormattingModel(this.formattingSettings);
    }

    // 🔑 Bookmark support methods
    public getBookmarkState(): any {
        // Only persist the preset name; ranges are recalculated live on restore.
        // Explicitly set dates to null to prevent Power BI from capturing old date values
        return {
            lastPreset: this.activePreset,
            isClearSelection: false, // Flag to indicate if this is a clear selection bookmark
            selectedMinDate: null, // Explicitly null to prevent Power BI from capturing old dates
            selectedMaxDate: null, // Explicitly null to prevent Power BI from capturing old dates
            dataMinDate: null, // Explicitly null to prevent Power BI from capturing old data bounds
            dataMaxDate: null  // Explicitly null to prevent Power BI from capturing old data bounds
        };
    }



    //         // Set flag to prevent update cycle from overriding restored state
    //         this.isRestoringBookmark = true;

    //         // CRITICAL EARLY CHECK: If bookmark has a time-based preset, recalculate IMMEDIATELY
    //         // This must happen BEFORE any dates are restored to prevent old dates from being set
    //         // Check multiple sources: state.lastPreset, this.activePreset, format pane
    //         const presetFromBookmark = state?.lastPreset;
    //         const presetFromActive = this.activePreset;
    //         const presetFromFormatPane = String(this.formattingSettings.presetsCard.preset.value);
    //         const detectedPreset = presetFromBookmark || presetFromActive || (presetFromFormatPane !== "none" ? presetFromFormatPane : null);

    //         if (detectedPreset && detectedPreset !== "none") {
    //             const isTimeBased = ["today", "yesterday", "last3days", "last7Days", "last30Days", "thisMonth", "lastMonth"].includes(detectedPreset);
    //             if (isTimeBased) {
    //                 console.log("EARLY CHECK: Time-based preset detected - recalculating BEFORE any date restore:", {
    //                     preset: detectedPreset,
    //                     source: presetFromBookmark ? "bookmark" : presetFromActive ? "activePreset" : "formatPane",
    //                     storedDates: {
    //                         min: state?.selectedMinDate,
    //                         max: state?.selectedMaxDate
    //                     },
    //                     willRecalculate: true,
    //                     currentDate: new Date().toISOString()
    //                 });
    //                 this.lastPresetSetting = detectedPreset;
    //                 this.activePreset = detectedPreset;

    //                 // CRITICAL: Reset manual selection flag - clear filter means we're going back to preset
    //                 this.hasManualSelection = false;

    //                 // CRITICAL: Calculate preset range IMMEDIATELY and update selectedMinDate/selectedMaxDate
    //                 // This must happen BEFORE any rendering to ensure correct dates are used
    //                 const presetRange = this.calculatePresetRange(detectedPreset);
    //                 if (presetRange) {
    //                     this.selectedMinDate = presetRange.from;
    //                     this.selectedMaxDate = presetRange.to;
    //                     console.log("IMMEDIATELY updated selectedMinDate/selectedMaxDate with recalculated preset:", {
    //                         preset: detectedPreset,
    //                         newMin: this.selectedMinDate.toISOString(),
    //                         newMax: this.selectedMaxDate.toISOString(),
    //                         hasManualSelection: this.hasManualSelection
    //                     });
    //                 }

    //                 // Also call clearSelectionWithCurrentData to ensure everything is consistent
    //                 this.clearSelectionWithCurrentData();

    //                 // CRITICAL: Ensure filter is applied to other visuals
    //                 // clearSelectionWithCurrentData applies filter, but if currentDataSource wasn't available,
    //                 // we need to ensure it gets applied when update() runs
    //                 // if (this.currentDataSource && this.selectedMinDate && this.selectedMaxDate) {
    //                 //     console.log("Early check - applying filter to other visuals:", {
    //                 //         min: this.selectedMinDate.toISOString(),
    //                 //         max: this.selectedMaxDate.toISOString()
    //                 //     });
    //                 //     this.shouldApplyFilter = true;
    //                 //     this.isUserInitiatedChange = true;
    //                 //     this.applyDateFilter(this.currentDataSource, this.selectedMinDate, this.selectedMaxDate);
    //                 // } else {
    //                 //     console.log("Early check - data source not available yet, filter will be applied in update()");
    //                 //     // Mark that we need to apply filter when update() runs
    //                 //     this.shouldApplyFilter = true;
    //                 //     this.isUserInitiatedChange = true;
    //                 // }



    //                 let minToApply: Date | null = this.selectedMinDate;
    // let maxToApply: Date | null = this.selectedMaxDate;

    // // If a recalculated preset exists (for time-based preset), use it instead of stored dates
    // if (this.presetRangeForClear) {
    //     minToApply = this.presetRangeForClear.from;
    //     maxToApply = this.presetRangeForClear.to;

    //     console.log("Early check - using recalculated preset range for filter:", {
    //         min: minToApply.toISOString(),
    //         max: maxToApply.toISOString(),
    //         preset: this.activePreset
    //     });
    // }

    // if (this.currentDataSource && minToApply && maxToApply) {
    //     console.log("Early check - applying filter to other visuals:", {
    //         min: minToApply.toISOString(),
    //         max: maxToApply.toISOString()
    //     });

    //     this.shouldApplyFilter = true;
    //     this.isUserInitiatedChange = true;

    //     this.applyDateFilter(this.currentDataSource, minToApply, maxToApply);
    // } else {
    //     console.log("Early check - data source not available yet, filter will be applied in update()");
    //     this.shouldApplyFilter = true;
    //     this.isUserInitiatedChange = true;
    // }


    //                 // Update React component immediately with recalculated dates
    //                 if (this.reactSliderWrapper && this.selectedMinDate && this.selectedMaxDate) {
    //                     console.log("Updating React component with recalculated dates:", {
    //                         min: this.selectedMinDate.toISOString(),
    //                         max: this.selectedMaxDate.toISOString()
    //                     });
    //                     this.reactSliderWrapper.updateDates(this.selectedMinDate, this.selectedMaxDate);
    //                 }
    //                 setTimeout(() => {
    //                     this.isRestoringBookmark = false;
    //                     this.isUserInitiatedChange = false;
    //                 }, 300);
    //                 return; // Return early - never restore old dates
    //             }
    //         }

    //         // Check if this is a clear selection bookmark
    //         if (state && state.isClearSelection === true) {
    //             console.log("Clear selection bookmark detected - recalculating preset based on current data");
    //             // CRITICAL: Reset manual selection flag - clear filter means we're going back to preset
    //             this.hasManualSelection = false;

    //             // CRITICAL FIX: Set lastPresetSetting from bookmark state BEFORE calling clearSelectionWithCurrentData
    //             // This ensures clearSelectionWithCurrentData uses the correct preset from the bookmark
    //             if (state.lastPreset && state.lastPreset !== "none") {
    //                 this.lastPresetSetting = state.lastPreset;
    //                 this.activePreset = state.lastPreset;
    //                 console.log("Clear selection - preset from bookmark:", state.lastPreset);
    //             } else {
    //                 // Fallback to format pane preset if bookmark doesn't have preset
    //                 const formatPanePreset = String(this.formattingSettings.presetsCard.preset.value);
    //                 if (formatPanePreset && formatPanePreset !== "none") {
    //                     this.lastPresetSetting = formatPanePreset;
    //                     this.activePreset = formatPanePreset;
    //                     console.log("Clear selection - using format pane preset:", formatPanePreset);
    //                 }

    //             }
    //             // When clearing, always recalculate preset based on CURRENT date and CURRENT data bounds
    //             // This ensures that for live data, "last 3 days" means the latest 3 days, not old dates
    //             this.clearSelectionWithCurrentData();

    //             // CRITICAL: Ensure filter is applied to other visuals
    //             // clearSelectionWithCurrentData applies filter, but ensure it's applied here too
    //             // if (this.currentDataSource && this.selectedMinDate && this.selectedMaxDate) {
    //             //     console.log("Clear selection - applying filter to other visuals:", {
    //             //         min: this.selectedMinDate.toISOString(),
    //             //         max: this.selectedMaxDate.toISOString()
    //             //     });
    //             //     this.shouldApplyFilter = true;
    //             //     this.isUserInitiatedChange = true;
    //             //     this.applyDateFilter(this.currentDataSource, this.selectedMinDate, this.selectedMaxDate);
    //             // } else {
    //             //     console.log("Clear selection - data source not available yet, filter will be applied in update()");
    //             //     // Mark that we need to apply filter when update() runs
    //             //     this.shouldApplyFilter = true;
    //             //     this.isUserInitiatedChange = true;
    //             // }


    //             let minToApply: Date | null = null;
    // let maxToApply: Date | null = null;

    // // If clear selection triggered a preset recalculation → use dynamic preset range
    // if (this.presetRangeForClear) {
    //     minToApply = this.presetRangeForClear.from;
    //     maxToApply = this.presetRangeForClear.to;

    //     console.log("Clear selection - using dynamic preset range:", {
    //         min: minToApply.toISOString(),
    //         max: maxToApply.toISOString(),
    //         preset: this.activePreset
    //     });
    // } else {
    //     // fallback (manual selection or non-preset)
    //     minToApply = this.selectedMinDate;
    //     maxToApply = this.selectedMaxDate;
    // }

    // if (this.currentDataSource && minToApply && maxToApply) {
    //     console.log("Clear selection - applying filter to other visuals:", {
    //         min: minToApply.toISOString(),
    //         max: maxToApply.toISOString()
    //     });

    //     this.shouldApplyFilter = true;
    //     this.isUserInitiatedChange = true;

    //     this.applyDateFilter(this.currentDataSource, minToApply, maxToApply);
    // } else {
    //     console.log("Clear selection - data source not available yet, filter will be applied in update()");
    //     this.shouldApplyFilter = true;
    //     this.isUserInitiatedChange = true;
    // }


    //             // Update React component immediately with recalculated dates
    //             if (this.reactSliderWrapper && this.selectedMinDate && this.selectedMaxDate) {
    //                 console.log("Clear selection - updating React component with recalculated dates:", {
    //                     min: this.selectedMinDate.toISOString(),
    //                     max: this.selectedMaxDate.toISOString(),
    //                     hasManualSelection: this.hasManualSelection
    //                 });
    //                 this.reactSliderWrapper.updateDates(this.selectedMinDate, this.selectedMaxDate);
    //             }
    //             setTimeout(() => {
    //                 this.isRestoringBookmark = false;
    //                 this.isUserInitiatedChange = false;
    //             }, 300);
    //             return;
    //         }

    //         // Check if recalculate preset on bookmark is enabled
    //         const recalculatePreset = this.formattingSettings.presetsCard.recalculatePresetOnBookmark.value;
    //         const bookmarkState = state;

    //         // CRITICAL FIX: For time-based presets, ALWAYS recalculate based on CURRENT date
    //         // This ensures "yesterday" means today's yesterday, not the date from when bookmark was created

    //         // First, try to get preset from bookmark state
    //         let presetToUse: string | null = null;
    //         if (bookmarkState && bookmarkState.lastPreset && bookmarkState.lastPreset !== "none") {
    //             presetToUse = bookmarkState.lastPreset;
    //             console.log("Bookmark restore - preset from bookmark state:", presetToUse);
    //         } else {
    //             // Fallback: Check if current format pane has a time-based preset set
    //             // This handles cases where bookmark state might not have stored the preset correctly
    //             const currentFormatPanePreset = String(this.formattingSettings.presetsCard.preset.value);
    //             if (currentFormatPanePreset && currentFormatPanePreset !== "none") {
    //                 const isTimeBased = ["today", "yesterday", "last3days", "last7Days", "last30Days", "thisMonth", "lastMonth"].includes(currentFormatPanePreset);
    //                 if (isTimeBased) {
    //                     presetToUse = currentFormatPanePreset;
    //                     console.log("Bookmark restore - using format pane preset as fallback:", presetToUse);
    //                 }
    //             }
    //         }

    //         // If we have a preset, check if it's time-based and should be recalculated
    //         if (presetToUse) {
    //             const isTimeBasedPreset = ["today", "yesterday", "last3days", "last7Days", "last30Days", "thisMonth", "lastMonth"].includes(presetToUse);

    //             console.log("Bookmark restore - preset analysis:", {
    //                 preset: presetToUse,
    //                 isTimeBasedPreset: isTimeBasedPreset,
    //                 recalculateEnabled: recalculatePreset,
    //                 willRecalculate: isTimeBasedPreset || recalculatePreset
    //             });

    //             // For time-based presets, ALWAYS recalculate (they should be dynamic)
    //             // For data-bound presets (minDate, maxDate), respect the recalculatePresetOnBookmark setting
    //             if (isTimeBasedPreset || recalculatePreset) {
    //                 console.log("Recalculating preset on bookmark restore:", {
    //                     preset: presetToUse,
    //                     isTimeBasedPreset: isTimeBasedPreset,
    //                     recalculateEnabled: recalculatePreset,
    //                     currentDate: new Date().toISOString(),
    //                     storedMinDate: bookmarkState?.selectedMinDate,
    //                     storedMaxDate: bookmarkState?.selectedMaxDate
    //                 });

    //                 // Restore the preset setting so clearSelectionWithCurrentData uses it
    //                 this.activePreset = presetToUse;
    //                 this.lastPresetSetting = presetToUse;

    //                 // IMPORTANT: Don't restore the old dates - let clearSelectionWithCurrentData calculate fresh dates
    //                 // This ensures "yesterday" means today's yesterday, not the date from when bookmark was created

    //                 // Call clearSelectionWithCurrentData which will recalculate the preset based on CURRENT data
    //                 this.clearSelectionWithCurrentData();

    //                 setTimeout(() => {
    //                     this.isRestoringBookmark = false;
    //                 }, 300);
    //                 return; // CRITICAL: Return here to prevent restoring old dates below
    //             } else {
    //                 console.log("Preset found but not recalculating (data-bound preset with recalculate disabled):", {
    //                     preset: presetToUse
    //                 });
    //             }
    //         } else {
    //             console.log("No preset found in bookmark state or format pane:", {
    //                 hasBookmarkState: !!bookmarkState,
    //                 bookmarkLastPreset: bookmarkState?.lastPreset,
    //                 formatPanePreset: String(this.formattingSettings.presetsCard.preset.value)
    //             });

    //             // SAFEGUARD: Even if preset wasn't detected above, check format pane for time-based preset
    //             // This ensures we don't restore old dates for time-based presets
    //             const formatPanePreset = String(this.formattingSettings.presetsCard.preset.value);
    //             if (formatPanePreset && formatPanePreset !== "none") {
    //                 const isTimeBased = ["today", "yesterday", "last3days", "last7Days", "last30Days", "thisMonth", "lastMonth"].includes(formatPanePreset);
    //                 if (isTimeBased) {
    //                     console.log("Safeguard: Format pane has time-based preset, recalculating:", formatPanePreset);
    //                     this.lastPresetSetting = formatPanePreset;
    //                     this.activePreset = formatPanePreset;
    //                     this.clearSelectionWithCurrentData();
    //                     setTimeout(() => {
    //                         this.isRestoringBookmark = false;
    //                     }, 300);
    //                     return; // Don't restore old dates
    //                 }
    //             }
    //         }


    //         // CRITICAL SAFEGUARD: Before restoring ANY dates, check if bookmark has a time-based preset
    //         // If it does, we MUST recalculate instead of restoring old dates
    //         // This ensures that bookmarks with preset names always use current date, not stored dates
    //         if (bookmarkState?.lastPreset && bookmarkState.lastPreset !== "none") {
    //             const isTimeBasedPreset = ["today", "yesterday", "last3days", "last7Days", "last30Days", "thisMonth", "lastMonth"].includes(bookmarkState.lastPreset);
    //             if (isTimeBasedPreset) {
    //                 console.log("CRITICAL: Bookmark has time-based preset - IGNORING stored dates and recalculating:", {
    //                     preset: bookmarkState.lastPreset,
    //                     storedDates: {
    //                         min: bookmarkState.selectedMinDate,
    //                         max: bookmarkState.selectedMaxDate
    //                     },
    //                     willRecalculate: true,
    //                     currentDate: new Date().toISOString()
    //                 });
    //                 this.lastPresetSetting = bookmarkState.lastPreset;
    //                 this.activePreset = bookmarkState.lastPreset;
    //                 this.clearSelectionWithCurrentData();
    //                 setTimeout(() => {
    //                     this.isRestoringBookmark = false;
    //                 }, 300);
    //                 return; // NEVER restore old dates for time-based presets - always recalculate
    //             }
    //         }

    //         // Restore dates from bookmark state
    //         // NOTE: This will only execute if preset is NOT time-based or preset recalculation was skipped
    //         // For time-based presets, we should have returned above
    //         if (bookmarkState.selectedMinDate) {
    //             this.selectedMinDate = new Date(bookmarkState.selectedMinDate);
    //         }
    //         if (bookmarkState.selectedMaxDate) {
    //             this.selectedMaxDate = new Date(bookmarkState.selectedMaxDate);
    //         }
    //         if (bookmarkState.dataMinDate) {
    //             this.dataMinDate = new Date(bookmarkState.dataMinDate);
    //         }
    //         if (bookmarkState.dataMaxDate) {
    //             this.dataMaxDate = new Date(bookmarkState.dataMaxDate);
    //         }
    //         if (bookmarkState.lastPreset) {
    //             this.activePreset = bookmarkState.lastPreset;
    //             this.lastPresetSetting = bookmarkState.lastPreset;
    //         }

    //         // Apply the restored filter if we have a data source
    //         // Note: currentDataSource might not be set yet if update hasn't run
    //         // The filter will be applied when update runs if currentDataSource becomes available
    //         // if (this.currentDataSource && this.selectedMinDate && this.selectedMaxDate) {
    //         //     this.shouldApplyFilter = true;
    //         //     this.isUserInitiatedChange = true;
    //         //     console.log("Applying filter from bookmark restore:", {
    //         //         min: this.selectedMinDate.toISOString(),
    //         //         max: this.selectedMaxDate.toISOString()
    //         //     });
    //         //     this.applyDateFilter(this.currentDataSource, this.selectedMinDate, this.selectedMaxDate);
    //         //     setTimeout(() => {
    //         //         this.isUserInitiatedChange = false;
    //         //     }, 100);
    //         // } else {
    //         //     // Data source not available yet - mark that we need to apply filter when update runs
    //         //     console.log("Bookmark restored but data source not available yet - will apply filter on next update");
    //         //     this.shouldApplyFilter = true;
    //         //     this.isUserInitiatedChange = true;
    //         // }


    //         let minToApply: Date | null = null;
    // let maxToApply: Date | null = null;

    // // If we came from a bookmark AND you have dynamic preset range recalculated
    // if (this.isRestoringBookmark && this.presetRangeForClear) {
    //     minToApply = this.presetRangeForClear.from;
    //     maxToApply = this.presetRangeForClear.to;

    //     console.log("Using dynamic preset range after bookmark restore:", {
    //         min: minToApply.toISOString(),
    //         max: maxToApply.toISOString(),
    //         preset: this.activePreset
    //     });
    // } else {
    //     // Normal behavior (manual selection or normal slider change)
    //     minToApply = this.selectedMinDate;
    //     maxToApply = this.selectedMaxDate;
    // }

    // if (this.currentDataSource && minToApply && maxToApply) {
    //     this.shouldApplyFilter = true;
    //     this.isUserInitiatedChange = true;

    //     console.log("Applying filter:", {
    //         min: minToApply.toISOString(),
    //         max: maxToApply.toISOString()
    //     });

    //     this.applyDateFilter(this.currentDataSource, minToApply, maxToApply);

    //     setTimeout(() => {
    //         this.isUserInitiatedChange = false;
    //     }, 100);
    // } else {
    //     console.log("Data source not ready – will apply filter on next update");
    //     this.shouldApplyFilter = true;
    //     this.isUserInitiatedChange = true;
    // }


    //         // Update React component if it exists
    //         if (this.reactSliderWrapper && this.selectedMinDate && this.selectedMaxDate) {
    //             this.reactSliderWrapper.updateDates(this.selectedMinDate, this.selectedMaxDate);
    //         }

    //         // Force re-render to show restored state
    //         this.needsReRender = true;

    //         // Safety timeout to clear the flag if update cycle doesn't run
    //         // The update cycle will clear it earlier if it applies the filter
    //         setTimeout(() => {
    //             if (this.isRestoringBookmark) {
    //                 this.isRestoringBookmark = false;
    //                 this.isUserInitiatedChange = false;
    //                 console.log("Bookmark restore flag cleared by safety timeout");
    //             }
    //         }, 500);
    //     }








    // 🔑 Clear selection method for bookmark interaction


    public restoreBookmarkState(state: any): void {
        console.log("Restoring bookmark state:", state);
        console.log("WARNING: Power BI may have captured old dates. We will recalculate based on current date for time-based presets.");

        this.isRestoringBookmark = true;

        // If this restore came from an external "clear filters" bookmark/button,
        // immediately recalc the preset against current date/data and stop.
        // This ensures clear button always uses CURRENT date, not bookmark creation date
        if (state?.isClearSelection) {
            console.log("Clear selection bookmark detected - will use CURRENT date, not bookmark creation date");
            
            // Prefer the preset captured in the bookmark; otherwise fall back to the current format pane preset
            const clearPreset = (state.lastPreset && state.lastPreset !== "none")
                ? state.lastPreset
                : String(this.formattingSettings.presetsCard.preset.value);

            if (clearPreset && clearPreset !== "none") {
                this.lastPresetSetting = clearPreset;
                this.activePreset = clearPreset;
            }

            // Always recalc using live date/data; applies filter to other visuals if possible
            // This ensures we use TODAY's date, not the date when bookmark was created
            this.clearSelectionWithCurrentData();

            // Ensure flags are cleared so subsequent updates behave normally
            setTimeout(() => {
                this.isRestoringBookmark = false;
                this.isUserInitiatedChange = false;
            }, 150);

            return;
        }

        // Determine preset from bookmark / active / format pane
        const presetFromBookmark = state?.lastPreset;
        const presetFromActive = this.activePreset;
        const presetFromFormatPane = String(this.formattingSettings.presetsCard.preset.value);
        const detectedPreset = presetFromBookmark || presetFromActive || (presetFromFormatPane !== "none" ? presetFromFormatPane : null);


        console.log("presetFromBookmark", presetFromBookmark)
        console.log("presetFromActive ", presetFromActive)
        console.log("presetFromActive ", presetFromActive)
        console.log("presetFromActive ", presetFromActive)

        // Check if time-based preset
        const timeBasedPresets = ["today", "yesterday", "last3days", "last7Days", "last30Days", "thisMonth", "lastMonth"];
        const isTimeBased = detectedPreset && timeBasedPresets.includes(detectedPreset);

        // CRITICAL: For time-based presets, ALWAYS recalculate based on current date
        // NEVER use dates from bookmark state - Power BI may have captured old dates
        if (isTimeBased) {
            console.log("Time-based preset detected - recalculating based on CURRENT date:", detectedPreset);
            console.log("IGNORING Power BI captured dates (will use live calculation):", {
                capturedMinDate: state?.selectedMinDate,
                capturedMaxDate: state?.selectedMaxDate
            });

            // Recalculate preset range dynamically (always live date)
            // CRITICAL: Never use old dates from bookmark state for time-based presets
            const presetRange = this.calculatePresetRange(detectedPreset);
            if (presetRange) {
                // Force use of recalculated dates - ignore any dates Power BI might have stored
                this.selectedMinDate = presetRange.from;
                this.selectedMaxDate = presetRange.to;
                this.presetRangeForClear = presetRange;

                // Reset manual selection and record preset
                this.hasManualSelection = false;
                this.activePreset = detectedPreset;
                this.lastPresetSetting = detectedPreset;

                // Always mark for filter application; if data source is ready, apply now
                this.shouldApplyFilter = true;
                this.isUserInitiatedChange = true; // force apply in update path too

                if (this.currentDataSource) {
                    this.applyDateFilter(this.currentDataSource, presetRange.from, presetRange.to);
                } else {
                    console.log("Time-based preset restore: data source not ready, will apply in update()");
                    this.needsReRender = true;
                }

                // Update slider / React component immediately with recalculated dates
                if (this.reactSliderWrapper) {
                    this.reactSliderWrapper.updateDates(presetRange.from, presetRange.to, presetRange);
                }

                // Keep isRestoringBookmark true; update() will clear it after applying filter
                return; // never restore old bookmark dates - always use live calculation
            }
        }

        // For non-time-based presets, check if Power BI stored dates
        // If dates exist in state, they might be from Power BI's filter capture
        // Only use them if they're reasonable (not from bookmark creation day)
        if (state.selectedMinDate && state.selectedMaxDate) {
            const stateMinDate = new Date(state.selectedMinDate);
            const stateMaxDate = new Date(state.selectedMaxDate);
            
            // Check if these dates are likely from Power BI's filter capture (not our bookmark state)
            // If we have a preset, prefer recalculating it instead of using old dates
            if (detectedPreset && detectedPreset !== "none") {
                console.log("Non-time-based preset detected - recalculating instead of using Power BI captured dates:", detectedPreset);
                const presetRange = this.calculatePresetRange(detectedPreset);
                if (presetRange) {
                    this.selectedMinDate = presetRange.from;
                    this.selectedMaxDate = presetRange.to;
                    this.presetRangeForClear = presetRange;
                } else {
                    // Fallback to state dates if preset calculation fails
                    this.selectedMinDate = stateMinDate;
                    this.selectedMaxDate = stateMaxDate;
                }
            } else {
                // No preset - use dates from state
                this.selectedMinDate = stateMinDate;
                this.selectedMaxDate = stateMaxDate;
            }
        }
        
        // Restore data bounds if available
        if (state.dataMinDate) this.dataMinDate = new Date(state.dataMinDate);
        if (state.dataMaxDate) this.dataMaxDate = new Date(state.dataMaxDate);
        
        // Restore preset info
        if (state.lastPreset && state.lastPreset !== "none") {
            this.activePreset = state.lastPreset;
            this.lastPresetSetting = state.lastPreset;
        }

        // Apply restored filter
        if (this.currentDataSource && this.selectedMinDate && this.selectedMaxDate) {
            this.applyDateFilter(this.currentDataSource, this.selectedMinDate, this.selectedMaxDate);
            this.shouldApplyFilter = true;
            this.isUserInitiatedChange = true;
        }

        // Update React slider
        if (this.reactSliderWrapper && this.selectedMinDate && this.selectedMaxDate) {
            this.reactSliderWrapper.updateDates(this.selectedMinDate, this.selectedMaxDate);
        }

        setTimeout(() => {
            this.isRestoringBookmark = false;
            this.isUserInitiatedChange = false;
        }, 300);
    }






    public clearSelection(): void {
        // Prefer the last preset (bookmark/format pane) and recalc live range
        const selectedPreset = this.lastPresetSetting || String(this.formattingSettings.presetsCard.preset.value);
        const presetRange = (selectedPreset && selectedPreset !== "none") ? this.calculatePresetRange(selectedPreset) : null;

        if (presetRange) {
            this.selectedMinDate = presetRange.from;
            this.selectedMaxDate = presetRange.to;
            this.activePreset = selectedPreset;
            this.lastPresetSetting = selectedPreset;
            this.hasManualSelection = false;

            console.log("clearSelection - live preset applied:", {
                preset: selectedPreset,
                selectedMinDate: this.selectedMinDate.toISOString(),
                selectedMaxDate: this.selectedMaxDate.toISOString(),
                presetRangeFrom: presetRange.from.toISOString(),
                presetRangeTo: presetRange.to.toISOString()
            });
        } else {
            // No preset - use data bounds
            this.selectedMinDate = this.dataMinDate;
            this.selectedMaxDate = this.dataMaxDate;
            this.activePreset = null;
            this.hasManualSelection = false;
        }

        // Apply the filter using the live preset range (or bounds)
        if (this.currentDataSource && this.selectedMinDate && this.selectedMaxDate) {
            this.shouldApplyFilter = true;
            this.isUserInitiatedChange = true;
            this.applyDateFilter(this.currentDataSource, this.selectedMinDate, this.selectedMaxDate);
            setTimeout(() => {
                this.isUserInitiatedChange = false;
            }, 100);
        } else {
            // Defer application to update()
            this.shouldApplyFilter = true;
            this.isUserInitiatedChange = true;
            this.needsReRender = true;
        }

        // Force re-render
        this.needsReRender = true;

        // Update React component immediately to sync currentMinDate and currentMaxDate with presetRange
        if (this.reactSliderWrapper && this.selectedMinDate && this.selectedMaxDate) {
            // Store the presetRange for use in updateDates
            this.presetRangeForClear = presetRange;
            this.reactSliderWrapper.updateDates(this.selectedMinDate, this.selectedMaxDate, presetRange);
        }
    }

    // 🔑 Clear selection method that always uses CURRENT data (for live data scenarios)
    // This ensures that when clearing filters, ALL presets are recalculated based on:
    // - Current date (for time-based presets: today, yesterday, last3days, last7Days, last30Days, thisMonth, lastMonth)
    // - Current data bounds (for data-bound presets: minDate, maxDate)
    // This prevents using old bookmark dates when new live data arrives
    private clearSelectionWithCurrentData(): void {
        console.log("Clearing selection with current data - recalculating preset based on latest dates/data");

        // Use the preset from lastPresetSetting if available (from bookmark), 
        // otherwise use the CURRENT preset setting from format pane
        // This ensures that when restoring a bookmark with a preset, we use the bookmark's preset
        // but recalculate it based on CURRENT date/data
        const selectedPreset = this.lastPresetSetting || String(this.formattingSettings.presetsCard.preset.value);

        console.log("clearSelectionWithCurrentData - preset source:", {
            usingBookmarkPreset: !!this.lastPresetSetting,
            bookmarkPreset: this.lastPresetSetting,
            formatPanePreset: String(this.formattingSettings.presetsCard.preset.value),
            selectedPreset: selectedPreset,
            currentDate: new Date().toISOString()
        });

        // Recalculate preset range - this will use:
        // - Current date (new Date()) for time-based presets (today, yesterday, last3days, etc.)
        // - Current data bounds (this.dataMinDate/this.dataMaxDate) for data-bound presets (minDate, maxDate)
        // - Clamping logic ensures all presets respect current data bounds
        const presetRange = (selectedPreset && selectedPreset !== "none") ? this.calculatePresetRange(selectedPreset) : null;

        if (presetRange) {
            // Use the recalculated range for all preset types
            // calculatePresetRange already handles:
            // - Time-based presets: uses current date (new Date())
            // - Data-bound presets: uses current data bounds (this.dataMinDate/this.dataMaxDate)
            // - Clamping: ensures range respects current data bounds
            this.selectedMinDate = presetRange.from;
            this.selectedMaxDate = presetRange.to;

            // Update activePreset tracking
            this.activePreset = selectedPreset;
            this.lastPresetSetting = selectedPreset;
            // Clearing preset selection removes any manual override
            this.hasManualSelection = false;

            console.log("Preset recalculated for clear filter (all preset types supported):", {
                preset: selectedPreset,
                minDate: this.selectedMinDate.toISOString(),
                maxDate: this.selectedMaxDate.toISOString(),
                currentDataMin: this.dataMinDate?.toISOString(),
                currentDataMax: this.dataMaxDate?.toISOString()
            });
        } else {
            // No preset - use CURRENT data bounds (which should reflect latest live data)
            this.selectedMinDate = this.dataMinDate;
            this.selectedMaxDate = this.dataMaxDate;
            this.activePreset = null;
            this.hasManualSelection = false;

            console.log("No preset - using current data bounds:", {
                minDate: this.selectedMinDate?.toISOString(),
                maxDate: this.selectedMaxDate?.toISOString()
            });
        }

        // Apply the filter when clearing selection (user-initiated action)
        // if (this.currentDataSource && this.selectedMinDate && this.selectedMaxDate) {
        //     this.shouldApplyFilter = true;
        //     this.isUserInitiatedChange = true;
        //     this.applyDateFilter(this.currentDataSource, this.selectedMinDate, this.selectedMaxDate);
        //     setTimeout(() => {
        //         this.isUserInitiatedChange = false;
        //     }, 100);
        // }



        if (this.currentDataSource) {
            // Choose dynamic preset first (live date), fallback to manual selection
            const minDateToApply =
                this.presetRangeForClear?.from && this.isRestoringBookmark
                    ? this.presetRangeForClear.from
                    : this.selectedMinDate;

            const maxDateToApply =
                this.presetRangeForClear?.to && this.isRestoringBookmark
                    ? this.presetRangeForClear.to
                    : this.selectedMaxDate;

            if (minDateToApply && maxDateToApply) {
                this.shouldApplyFilter = true;
                this.isUserInitiatedChange = true;

                this.applyDateFilter(
                    this.currentDataSource,
                    minDateToApply,
                    maxDateToApply
                );

                setTimeout(() => {
                    this.isUserInitiatedChange = false;
                }, 100);
            }
        }


        // Force re-render
        this.needsReRender = true;

        // Update React component if it exists
        if (this.reactSliderWrapper && this.selectedMinDate && this.selectedMaxDate) {
            this.reactSliderWrapper.updateDates(this.selectedMinDate, this.selectedMaxDate);
        }
    }

    // 🔑 Alternative clear method for bookmark buttons
    public resetToDefault(): void {
        this.clearSelection();
    }

    public destroy(): void {
        if (this.boundContextMenuHandler) {
            this.target.removeEventListener("contextmenu", this.boundContextMenuHandler);
            this.boundContextMenuHandler = undefined;
        }

        if (this.reactSliderWrapper) {
            this.reactSliderWrapper.destroy();
            this.reactSliderWrapper = null;
        }
    }
}

