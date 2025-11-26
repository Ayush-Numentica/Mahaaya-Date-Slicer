"use strict";

import powerbi from "powerbi-visuals-api";
import { FormattingSettingsService } from "powerbi-visuals-utils-formattingmodel";
import "./../style/visual.less";

import VisualConstructorOptions = powerbi.extensibility.visual.VisualConstructorOptions;
import VisualUpdateOptions = powerbi.extensibility.visual.VisualUpdateOptions;
import IVisual = powerbi.extensibility.visual.IVisual;
import DialogAction = powerbi.DialogAction;

import { VisualFormattingSettingsModel } from "./settings";
import { ReactSliderWrapper } from "./ReactWrapper";
import { DatePickerDialog, DatePickerDialogResult } from "./DatePickerDialog";

export class Visual implements IVisual {
    private target: HTMLElement;
    private formattingSettings: VisualFormattingSettingsModel;
    private formattingSettingsService: FormattingSettingsService;
    private host: powerbi.extensibility.visual.IVisualHost;  // 🔑 keep host reference

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
    
    // Flag to track if user has manually selected dates (prevents preset from overriding manual selection)
    private hasManualSelection: boolean = false;

    constructor(options: VisualConstructorOptions) {
        this.target = options.element;
        this.host = options.host;   // ✅ store host
        this.formattingSettingsService = new FormattingSettingsService();
        this.formattingSettings = new VisualFormattingSettingsModel();
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

                    //  Apply filter to report (derive table/column from bound field)
                    if (this.dataMinDate && this.dataMaxDate) {
                        // Store the data source for universal filtering
                        this.currentDataSource = category.source;

                        const incomingFilters = (options.jsonFilters as powerbi.IFilter[]) || [];
                        const externalFilterRange = this.getDateRangeFromFilters(incomingFilters, category.source);
                        const filtersCleared = !externalFilterRange;

                        if (externalFilterRange && !this.isUserInitiatedChange) {
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
                        } else if (
                            !this.isUserInitiatedChange &&
                            filtersCleared &&
                            this.dataMinDate &&
                            this.dataMaxDate
                        ) {
                            // Filters cleared entirely – reset to preset range if available, otherwise data bounds
                            // Use current preset from formattingSettings instead of activePreset (which might be stale)
                            const currentPreset = String(this.formattingSettings.presetsCard.preset.value);
                            const presetRange = (currentPreset && currentPreset !== "none") ? this.calculatePresetRange(currentPreset) : null;
                            
                            if (presetRange) {
                                const resetMinDate = presetRange.from;
                                const resetMaxDate = presetRange.to;
                                
                                const needsReset = !this.selectedMinDate || 
                                                   !this.selectedMaxDate ||
                                                   this.selectedMinDate.getTime() !== resetMinDate.getTime() ||
                                                   this.selectedMaxDate.getTime() !== resetMaxDate.getTime();
                                
                                if (needsReset) {
                                    console.log("Filters cleared externally - reapplying preset range:", {
                                        preset: currentPreset,
                                        min: resetMinDate.toISOString(),
                                        max: resetMaxDate.toISOString()
                                    });
                                    
                                    // Update activePreset to match current setting
                                    this.activePreset = currentPreset;
                                    this.lastPresetSetting = currentPreset;
                                    this.hasManualSelection = false;
                                    
                                    // Treat as user action so the filter reapplies immediately
                                    this.isUserInitiatedChange = true;
                                    this.shouldApplyFilter = true;
                                    this.updateSelectedDates(resetMinDate, resetMaxDate, category?.source, true);
                                }
                            } else {
                                // No preset selected - snap back to full data bounds without reapplying filter
                                const resetMinDate = this.dataMinDate;
                                const resetMaxDate = this.dataMaxDate;
                                const needsReset = !this.selectedMinDate || 
                                                   !this.selectedMaxDate ||
                                                   this.selectedMinDate.getTime() !== resetMinDate.getTime() ||
                                                   this.selectedMaxDate.getTime() !== resetMaxDate.getTime();
                                
                                if (needsReset) {
                                    console.log("Filters cleared externally - resetting to data bounds (no active preset)");
                                    this.selectedMinDate = resetMinDate;
                                    this.selectedMaxDate = resetMaxDate;
                                    this.shouldApplyFilter = false;
                                    this.needsReRender = true;
                                }
                            }
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
                        
                        // Apply filter if conditions are met OR if preset just changed (force apply)
                        if (shouldApply || forceApplyForPreset) {
                            const filterMinDate = this.selectedMinDate;
                            const filterMaxDate = this.selectedMaxDate;
                            // Ensure filter is enabled when applying from update method
                            this.shouldApplyFilter = true;
                            console.log("Applying filter from update method:", {
                                isUserInitiated: this.isUserInitiatedChange,
                                isRestoringBookmark: this.isRestoringBookmark,
                                presetJustChanged: this.presetJustChanged,
                                forceApplyForPreset: forceApplyForPreset,
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

        const cardColor = (this.formattingSettings.dataPointCard.cardColor?.value?.value as string) || "#ffffff";

        this.target.innerHTML = "";

        const card = document.createElement("div");
        card.style.border = "1px solid #ccc";
        card.style.borderRadius = "8px";
        card.style.paddingTop="4px"
        card.style.paddingBottom="4px"
        // card.style.padding="0px"
        card.style.textAlign = "center";
        card.style.background = cardColor;
        card.style.fontFamily = "Segoe UI, sans-serif";


        const slicerHeader = !!this.formattingSettings.presetsCard.toShowHeader.value;

        if (slicerHeader) {
            const title = document.createElement("div");
            title.textContent = "Mahaaya Super Date Slicer";
            title.style.fontWeight = "bold";
            title.style.marginBottom = "5px";
            card.appendChild(title);
        }

        // Get formatting pane values
        const selectedStyle = this.formattingSettings.presetsCard.selectionStyle.value;
        const popupOnly = !!this.formattingSettings.presetsCard.toggleOption.value;

        if (popupOnly) {
            // Only show the dialog button when toggle is ON
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
            // Render slider or inline calendar, based on selection style
            const rangeSliderContainer = document.createElement("div");
            rangeSliderContainer.className = "range-slider-container";

            if (this.reactSliderWrapper) {
                this.reactSliderWrapper.destroy();
            }

            this.reactSliderWrapper = new ReactSliderWrapper(rangeSliderContainer);

            const inputFontSize = Number(this.formattingSettings.dataPointCard.fontSize.value) || 18;
            const inputFontColor = (this.formattingSettings.dataPointCard.fontColor?.value?.value as string) || "#000000";
            const inputBoxColor = (this.formattingSettings.dataPointCard.dateBoxColor?.value?.value as string) || "#ffffff";
            
            

            this.reactSliderWrapper.render({
                minDate,
                maxDate,
                currentMinDate: this.selectedMinDate || minDate,
                currentMaxDate: this.selectedMaxDate || maxDate,
                onDateChange: (newMinDate: Date, newMaxDate: Date) => {
                    this.updateSelectedDates(newMinDate, newMaxDate, source);
                },
                formatDateForSlider: this.formatDateForSlider.bind(this),
                onOpenDialog: this.openDatePickerDialog.bind(this),
                datePickerType: selectedStyle,
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
        return {
            selectedMinDate: this.selectedMinDate?.toISOString() || null,
            selectedMaxDate: this.selectedMaxDate?.toISOString() || null,
            dataMinDate: this.dataMinDate?.toISOString() || null,
            dataMaxDate: this.dataMaxDate?.toISOString() || null,
            lastPreset: this.activePreset,
            isClearSelection: false // Flag to indicate if this is a clear selection bookmark
        };
    }

    public restoreBookmarkState(state: any): void {
        console.log("Restoring bookmark state:", state);
        
        // Set flag to prevent update cycle from overriding restored state
        this.isRestoringBookmark = true;
        
        // Check if this is a clear selection bookmark
        if (state && state.isClearSelection === true) {
            console.log("Clear selection bookmark detected");
            this.clearSelection();
            setTimeout(() => {
                this.isRestoringBookmark = false;
            }, 300);
            return;
        }
        
        const bookmarkState = state;
        
        // Restore dates from bookmark state
        if (bookmarkState.selectedMinDate) {
            this.selectedMinDate = new Date(bookmarkState.selectedMinDate);
        }
        if (bookmarkState.selectedMaxDate) {
            this.selectedMaxDate = new Date(bookmarkState.selectedMaxDate);
        }
        if (bookmarkState.dataMinDate) {
            this.dataMinDate = new Date(bookmarkState.dataMinDate);
        }
        if (bookmarkState.dataMaxDate) {
            this.dataMaxDate = new Date(bookmarkState.dataMaxDate);
        }
        if (bookmarkState.lastPreset) {
            this.activePreset = bookmarkState.lastPreset;
            this.lastPresetSetting = bookmarkState.lastPreset;
        }

        // Apply the restored filter if we have a data source
        // Note: currentDataSource might not be set yet if update hasn't run
        // The filter will be applied when update runs if currentDataSource becomes available
        if (this.currentDataSource && this.selectedMinDate && this.selectedMaxDate) {
            this.shouldApplyFilter = true;
            this.isUserInitiatedChange = true;
            console.log("Applying filter from bookmark restore:", {
                min: this.selectedMinDate.toISOString(),
                max: this.selectedMaxDate.toISOString()
            });
            this.applyDateFilter(this.currentDataSource, this.selectedMinDate, this.selectedMaxDate);
            setTimeout(() => {
                this.isUserInitiatedChange = false;
            }, 100);
        } else {
            // Data source not available yet - mark that we need to apply filter when update runs
            console.log("Bookmark restored but data source not available yet - will apply filter on next update");
            this.shouldApplyFilter = true;
            this.isUserInitiatedChange = true;
        }

        // Update React component if it exists
        if (this.reactSliderWrapper && this.selectedMinDate && this.selectedMaxDate) {
            this.reactSliderWrapper.updateDates(this.selectedMinDate, this.selectedMaxDate);
        }

        // Force re-render to show restored state
        this.needsReRender = true;
        
        // Safety timeout to clear the flag if update cycle doesn't run
        // The update cycle will clear it earlier if it applies the filter
        setTimeout(() => {
            if (this.isRestoringBookmark) {
                this.isRestoringBookmark = false;
                this.isUserInitiatedChange = false;
                console.log("Bookmark restore flag cleared by safety timeout");
            }
        }, 500);
    }

    // 🔑 Clear selection method for bookmark interaction
    public clearSelection(): void {
        // Reset to current preset if available, otherwise reset to data bounds
        const selectedPreset = String(this.formattingSettings.presetsCard.preset.value);
        const presetRange = (selectedPreset && selectedPreset !== "none") ? this.calculatePresetRange(selectedPreset) : null;
        
        if (presetRange) {
            this.selectedMinDate = presetRange.from;
            this.selectedMaxDate = presetRange.to;
            // Update activePreset tracking
            this.activePreset = selectedPreset;
            this.lastPresetSetting = selectedPreset;
            // Clearing preset selection removes any manual override
            this.hasManualSelection = false;
        } else {
            // No preset - use data bounds
            this.selectedMinDate = this.dataMinDate;
            this.selectedMaxDate = this.dataMaxDate;
            this.activePreset = null;
            this.hasManualSelection = false;
        }

        // Apply the filter when clearing selection (user-initiated action)
        if (this.currentDataSource && this.selectedMinDate && this.selectedMaxDate) {
            this.shouldApplyFilter = true;
            this.isUserInitiatedChange = true;
            this.applyDateFilter(this.currentDataSource, this.selectedMinDate, this.selectedMaxDate);
            setTimeout(() => {
                this.isUserInitiatedChange = false;
            }, 100);
        }

        // Force re-render
        this.needsReRender = true;
    }

    // 🔑 Alternative clear method for bookmark buttons
    public resetToDefault(): void {
        this.clearSelection();
    }
}
