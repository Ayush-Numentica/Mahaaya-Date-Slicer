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

    // Track the last date selection made BY THE USER through this visual
    
    private lastUserSelectedMin: Date | null = null;
    private lastUserSelectedMax: Date | null = null;


    // ============================================================================
    // 🔑 CLEAR ALL SLICERS DETECTION STATE
    // ============================================================================
    // Track previous filter state to detect when filters go from present → empty
    // This is the only reliable way to detect "Clear all slicers" button clicks
    // since Power BI doesn't expose a direct event for it.
    // ============================================================================
    private previousFiltersExisted: boolean = false; // Track if filters existed in previous update
    private previousFilterHash: string | null = null; // Hash of previous filter state for comparison
    private isInitialLoad: boolean = true; // Track first render to avoid false positives

    // ============================================================================
    // 🔑 SYNC SLICER STATE TRACKING
    // ============================================================================
    // Power BI destroys and recreates visuals on page navigation, so we need to
    // distinguish between:
    // - Initial render (first time visual is created)
    // - Sync update (external filter from another page)
    // - User-initiated change (user interacts with slider/calendar)
    // ============================================================================
    private lastAppliedFilterHash: string | null = null; // Hash of the filter we last applied
    private isInitialRender: boolean = true; // True on first update() call after construction
    private isSyncUpdateFlag: boolean = false; // True when receiving external filter from sync


    // constructor(options: VisualConstructorOptions) {
    //     this.target = options.element;
    //     this.host = options.host;   // ✅ store host
    //     this.formattingSettingsService = new FormattingSettingsService();
    //     this.formattingSettings = new VisualFormattingSettingsModel();
    //     this.selectionManager = this.host.createSelectionManager();
    //     // Right-click anywhere inside the visual to show the built-in context menu
    //     this.boundContextMenuHandler = (event: MouseEvent) => {
    //         event.preventDefault();
    //         this.selectionManager.showContextMenu(undefined, {
    //             x: event.clientX,
    //             y: event.clientY
    //         });
    //     };
    //     this.target.addEventListener("contextmenu", this.boundContextMenuHandler);

    // }



    constructor(options: VisualConstructorOptions) {
            this.target = options.element;
            this.host = options.host;
            this.formattingSettingsService = new FormattingSettingsService();
            this.formattingSettings = new VisualFormattingSettingsModel();
            this.selectionManager = this.host.createSelectionManager();
            
            // Right-click context menu handler
            this.boundContextMenuHandler = (event: MouseEvent) => {
                event.preventDefault();
                this.selectionManager.showContextMenu(undefined, {
                    x: event.clientX,
                    y: event.clientY
                });
            };
            this.target.addEventListener("contextmenu", this.boundContextMenuHandler);

            // Initialize state for Clear All detection
            this.previousFiltersExisted = false;
            this.isInitialRender = true;
            
            // Initialize other state flags
            this.hasManualSelection = false;
            this.isUserInitiatedChange = false;
            this.isRestoringBookmark = false;
            this.presetJustChanged = false;
            this.shouldApplyFilter = false;
            this.needsReRender = false;
            
            // Initialize date state
            this.selectedMinDate = null;
            this.selectedMaxDate = null;
            this.dataMinDate = null;
            this.dataMaxDate = null;
            
            // Initialize preset state
            this.activePreset = null;
            this.lastPresetSetting = null;
            this.lastAppliedFilterHash = null;
            this.currentDataSource = null;
}



    public update(options: VisualUpdateOptions): void {
            



        console.log("New Version Loaded");

        const isService = typeof window !== 'undefined' && window.location && window.location.hostname.includes('app.powerbi.com');
        let detectedBookmarkRestore = false;

        const incomingFilters = (options.jsonFilters as powerbi.IFilter[]) || [];
        
        const wasInitialRender = this.isInitialRender;
        if (this.isInitialRender) {
            this.isInitialRender = false;
        }

        // Track filter state for Clear All detection
        const currentFiltersExist = incomingFilters.length > 0;
        
        // Detect Clear All: filters went from existing to empty
        const clearAllDetected = !wasInitialRender && 
            this.previousFiltersExisted && 
            !currentFiltersExist;
        console.log("clearAllDetected",clearAllDetected)
        // Update tracking for next cycle
        this.previousFiltersExisted = currentFiltersExist;


        console.log("=== UPDATE CYCLE ===");
        console.log("incomingFilters:", incomingFilters);
        console.log("incomingFilters.length:", incomingFilters.length);
        console.log("previousFiltersExisted:", this.previousFiltersExisted);
        console.log("wasInitialRender:", wasInitialRender);

        if (options.dataViews && options.dataViews[0]) {
            this.formattingSettings =
                this.formattingSettingsService.populateFormattingSettingsModel(
                    VisualFormattingSettingsModel,
                    options.dataViews[0]
                );

            const dataView = options.dataViews[0];
            const category = dataView.categorical?.categories?.[0];

            let presetRangeForClear: { from: Date; to: Date } | null = null;
            let defaultRangeForClear: { from: Date; to: Date } | null = null;

            if (category && category.values && category.values.length > 0) {
                if (category.source && this.isHierarchyField(category.source)) {
                    this.renderMessage("Please bind the base Date column, not the Date Hierarchy (Year/Quarter/Month/Day).From the Visual pane, right-click the Date field and select Date instead of Date Hierarchy.");
                    return;
                }

                const selectedPreset = String(this.formattingSettings.presetsCard.preset.value);
                const livePresetRange = this.calculatePresetRange(selectedPreset);

                this.currentDataSource = category.source;

                // Desktop: Check if data content changed
                if (!isService && this.dataMinDate && this.dataMaxDate) {
                    const currentDataHash = category.values.join('|');
                    if (this.lastDataHash && this.lastDataHash !== currentDataHash) {
                        this.needsReRender = true;
                    }
                    this.lastDataHash = currentDataHash;
                }

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

                    this.presetRangeForClear = presetRangeForClear;
                    this.defaultRangeForClear = defaultRangeForClear;

                    minDate.setHours(0, 0, 0, 0);
                    maxDate.setHours(23, 59, 59, 999);

                    const boundsChanged = (!this.dataMinDate || !this.dataMaxDate) ||
                        (this.dataMinDate.getTime() !== dataBoundMin.getTime() || this.dataMaxDate.getTime() !== dataBoundMax.getTime());

                    // Desktop: Force bounds change detection
                    if (!isService && !boundsChanged && this.dataMinDate && this.dataMaxDate) {
                        const timeDiff = Math.abs(this.dataMinDate.getTime() - minDate.getTime()) +
                            Math.abs(this.dataMaxDate.getTime() - maxDate.getTime());
                        if (timeDiff > 1000) {
                            this.dataMinDate = minDate;
                            this.dataMaxDate = maxDate;
                            this.selectedMinDate = minDate;
                            this.selectedMaxDate = maxDate;
                            this.needsReRender = true;
                        }
                    }

                    // Handle external slicer changes
                    if (boundsChanged && !this.isRestoringBookmark) {
                        const isInitialLoad = !this.dataMinDate || !this.dataMaxDate;
                        this.dataMinDate = minDate;
                        this.dataMaxDate = maxDate;

                        if (this.selectedMinDate && this.selectedMaxDate) {
                            const clampedMin = new Date(Math.max(this.selectedMinDate.getTime(), minDate.getTime()));
                            const clampedMax = new Date(Math.min(this.selectedMaxDate.getTime(), maxDate.getTime()));
                            this.selectedMinDate = clampedMin;
                            this.selectedMaxDate = clampedMax;
                        } else {
                            if (isInitialLoad && presetRangeForClear && selectedPreset && selectedPreset !== "none") {
                                this.selectedMinDate = presetRangeForClear.from;
                                this.selectedMaxDate = presetRangeForClear.to;
                                this.activePreset = selectedPreset;
                                this.lastPresetSetting = selectedPreset;

                                if (category?.source) {
                                    this.isUserInitiatedChange = true;
                                    this.shouldApplyFilter = true;
                                    this.updateSelectedDates(presetRangeForClear.from, presetRangeForClear.to, category.source, true);
                                    setTimeout(() => {
                                        this.isUserInitiatedChange = false;
                                    }, 100);
                                }
                            } else {
                                this.selectedMinDate = minDate;
                                this.selectedMaxDate = maxDate;
                            }
                        }
                        this.needsReRender = true;
                    } else if (boundsChanged && this.isRestoringBookmark) {
                        this.dataMinDate = minDate;
                        this.dataMaxDate = maxDate;
                    }

                    // Early external filter check
                    if (this.dataMinDate && this.dataMaxDate && category.source) {
                        this.currentDataSource = category.source;
                        const externalFilterRange = this.getDateRangeFromFilters(incomingFilters, category.source);
                        const externalFilterHash = this.createFilterHash(incomingFilters, category.source);

                        // Initial render with external filter
                        if (externalFilterRange && wasInitialRender && !this.isUserInitiatedChange && !this.isRestoringBookmark) {
                            console.log("🔄 EARLY INITIAL RENDER: Honoring external filter");
                            this.selectedMinDate = externalFilterRange.minDate;
                            this.selectedMaxDate = externalFilterRange.maxDate;
                            this.hasManualSelection = true;
                            this.lastAppliedFilterHash = externalFilterHash;
                        }
                    }

                    // Handle preset changes
                    const isPresetChange = this.lastPresetSetting !== selectedPreset;
                    const isInitialPresetLoad = !this.lastPresetSetting && selectedPreset && selectedPreset !== "none";

                    // FIX: When user changes preset, reset hasManualSelection to allow preset to apply
                    if (isPresetChange && !wasInitialRender) {
                        console.log("🔄 PRESET CHANGED: Resetting hasManualSelection to allow new preset");
                        this.hasManualSelection = false;
                    }

                    // Apply preset changes (now hasManualSelection won't block it)
                    if ((isPresetChange || isInitialPresetLoad) && !this.isRestoringBookmark && !this.hasManualSelection) {
                        this.dataMinDate = dataBoundMin;
                        this.dataMaxDate = dataBoundMax;

                        let presetSelectionMin: Date;
                        let presetSelectionMax: Date;

                        if (presetRangeForClear && selectedPreset && selectedPreset !== "none") {
                            presetSelectionMin = presetRangeForClear.from;
                            presetSelectionMax = presetRangeForClear.to;
                            this.activePreset = selectedPreset;
                        } else {
                            presetSelectionMin = dataBoundMin;
                            presetSelectionMax = dataBoundMax;
                            this.activePreset = null;
                        }

                        this.selectedMinDate = presetSelectionMin;
                        this.selectedMaxDate = presetSelectionMax;
                        this.lastPresetSetting = selectedPreset;
                        this.presetJustChanged = true;

                        this.isUserInitiatedChange = true;
                        this.shouldApplyFilter = true;
                        this.needsReRender = true;

                        if (category?.source) {
                            this.updateSelectedDates(presetSelectionMin, presetSelectionMax, category.source, true);
                            this.applyDateFilter(category.source, presetSelectionMin, presetSelectionMax);
                        }

                        setTimeout(() => {
                            this.isUserInitiatedChange = false;
                            this.presetJustChanged = false;
                        }, 500);
                    } else if (!boundsChanged) {
                        if (!this.dataMinDate || !this.dataMaxDate) {
                            this.dataMinDate = dataBoundMin;
                            this.dataMaxDate = dataBoundMax;
                        }

                        // Enforce preset if mismatch and no manual selection
                        if (presetRangeForClear && selectedPreset && selectedPreset !== "none" && !this.hasManualSelection) {
                            const presetMin = presetRangeForClear.from;
                            const presetMax = presetRangeForClear.to;

                            const selectionMismatch = !this.selectedMinDate ||
                                !this.selectedMaxDate ||
                                this.selectedMinDate.getTime() !== presetMin.getTime() ||
                                this.selectedMaxDate.getTime() !== presetMax.getTime();

                            if (selectionMismatch) {
                                this.selectedMinDate = presetMin;
                                this.selectedMaxDate = presetMax;
                                this.activePreset = selectedPreset;
                                this.lastPresetSetting = selectedPreset;
                                this.presetJustChanged = true;
                                this.isUserInitiatedChange = true;
                                this.shouldApplyFilter = true;

                                if (category?.source) {
                                    this.updateSelectedDates(presetMin, presetMax, category.source, true);
                                    this.applyDateFilter(category.source, presetMin, presetMax);
                                }

                                setTimeout(() => {
                                    this.isUserInitiatedChange = false;
                                    this.presetJustChanged = false;
                                }, 500);
                            } else {
                                this.activePreset = selectedPreset;
                                this.lastPresetSetting = selectedPreset;
                            }
                        } else if (!this.selectedMinDate || !this.selectedMaxDate) {
                            this.selectedMinDate = minDate;
                            this.selectedMaxDate = maxDate;
                        }
                    }

                    // Handle external filters
                    if (this.dataMinDate && this.dataMaxDate) {
                        this.currentDataSource = category.source;

                        const externalFilterRange = this.getDateRangeFromFilters(incomingFilters, category.source);
                        const externalFilterHash = this.createFilterHash(incomingFilters, category.source);
                        console.log("externalFilterRange",externalFilterRange);
                        console.log("externalFilterHash",externalFilterHash);

                        this.isSyncUpdateFlag = this.checkIsSyncUpdate(externalFilterHash);

                        // console.log("=== EXTERNAL FILTER DECISION ===");
                        // console.log("clearAllDetected:", clearAllDetected);
                        // console.log("externalFilterRange:", externalFilterRange);
                        // console.log("wasInitialRender:", wasInitialRender);
                        // console.log("isUserInitiatedChange:", this.isUserInitiatedChange);
                        // console.log("isRestoringBookmark:", this.isRestoringBookmark);
                        // console.log("isSyncUpdateFlag:", this.isSyncUpdateFlag);
                        // console.log("hasManualSelection:", this.hasManualSelection);
                        // console.log("lastUserSelectedMin:", this.lastUserSelectedMin);
                        // console.log("lastUserSelectedMax:", this.lastUserSelectedMax);

                        // FIX: Clear All Detection - Apply live preset dates
                        if (clearAllDetected) {
                            console.log("✅ CLEAR ALL DETECTED: Applying live preset dates");
                            
                            if (presetRangeForClear && selectedPreset && selectedPreset !== "none") {
                                // Apply live preset
                                this.selectedMinDate = presetRangeForClear.from;
                                this.selectedMaxDate = presetRangeForClear.to;
                                this.activePreset = selectedPreset;
                                this.lastPresetSetting = selectedPreset;
                                this.hasManualSelection = false;
                                
                                // Apply filter
                                this.applyDateFilter(category.source, presetRangeForClear.from, presetRangeForClear.to);
                                this.lastAppliedFilterHash = this.createFilterHash([{
                                    $schema: "http://powerbi.com/product/schema#advanced",
                                    target: { 
                                        table: this.extractTableAndColumn(category.source).tableName, 
                                        column: this.extractTableAndColumn(category.source).columnName 
                                    },
                                    logicalOperator: "And",
                                    conditions: [
                                        { operator: "GreaterThanOrEqual", value: presetRangeForClear.from.toISOString() },
                                        { operator: "LessThanOrEqual", value: presetRangeForClear.to.toISOString() }
                                    ]
                                } as powerbi.IFilter], category.source);
                                
                                this.needsReRender = true;
                                
                                console.log("✅ Applied live preset after Clear All:", {
                                    preset: selectedPreset,
                                    min: presetRangeForClear.from.toISOString(),
                                    max: presetRangeForClear.to.toISOString()
                                });
                            } else {
                                // No preset - use data bounds
                                this.selectedMinDate = minDate;
                                this.selectedMaxDate = maxDate;
                                this.hasManualSelection = false;
                                this.needsReRender = true;
                            }
                        }
                        // External filter on initial render
                        else if (externalFilterRange && wasInitialRender && !this.isUserInitiatedChange && !this.isRestoringBookmark) {
                            console.log("🔄 INITIAL RENDER WITH EXTERNAL FILTER: Honoring synced filter");

                            this.selectedMinDate = externalFilterRange.minDate;
                            this.selectedMaxDate = externalFilterRange.maxDate;
                            this.hasManualSelection = true;
                            this.shouldApplyFilter = false;
                            this.needsReRender = true;
                            this.lastAppliedFilterHash = externalFilterHash;

                            if (presetRangeForClear) {
                                const matchesPreset = Math.abs(externalFilterRange.minDate.getTime() - presetRangeForClear.from.getTime()) <= 1000 &&
                                    Math.abs(externalFilterRange.maxDate.getTime() - presetRangeForClear.to.getTime()) <= 1000;
                                if (matchesPreset && selectedPreset && selectedPreset !== "none") {
                                    this.activePreset = selectedPreset;
                                    this.lastPresetSetting = selectedPreset;
                                }
                            }
                            if (category?.source) {
                                    console.log("🔧 Applying filter to sync with other visuals");
                                    this.applyDateFilter(category.source, externalFilterRange.minDate, externalFilterRange.maxDate);
                                }
                            
                        }



                        
                        // Sync update 1
                        // else if (externalFilterRange && this.isSyncUpdateFlag && !this.isUserInitiatedChange && !this.isRestoringBookmark) {
                        //     console.log("🔄 SYNC UPDATE: Honoring external filter");

                        //     this.selectedMinDate = externalFilterRange.minDate;
                        //     this.selectedMaxDate = externalFilterRange.maxDate;
                        //     this.hasManualSelection = true;
                        //     this.shouldApplyFilter = false;
                        //     this.needsReRender = true;
                        //     this.lastAppliedFilterHash = externalFilterHash;

                        //     if (presetRangeForClear) {
                        //         const matchesPreset = Math.abs(externalFilterRange.minDate.getTime() - presetRangeForClear.from.getTime()) <= 1000 &&
                        //             Math.abs(externalFilterRange.maxDate.getTime() - presetRangeForClear.to.getTime()) <= 1000;
                        //         if (matchesPreset && selectedPreset && selectedPreset !== "none") {
                        //             this.activePreset = selectedPreset;
                        //             this.lastPresetSetting = selectedPreset;
                        //         }
                        //     }
                        // }

                        // Sync update 2
                        // else if (externalFilterRange && this.isSyncUpdateFlag && !this.isUserInitiatedChange && !this.isRestoringBookmark) {
                        //     console.log("🔄 SYNC UPDATE: Checking if matches live preset");
                            
                        //     // Check if external filter matches live preset
                        //     const matchesLivePreset = presetRangeForClear && 
                        //         Math.abs(externalFilterRange.minDate.getTime() - presetRangeForClear.from.getTime()) <= 1000 &&
                        //         Math.abs(externalFilterRange.maxDate.getTime() - presetRangeForClear.to.getTime()) <= 1000;
                            
                        //     // If external filter doesn't match live preset AND we have a preset set
                        //     if (!matchesLivePreset && presetRangeForClear && selectedPreset && selectedPreset !== "none") {
                        //         console.log("⚠️ External filter doesn't match live preset - applying live preset instead");
                                
                        //         // Apply live preset dates
                        //         this.selectedMinDate = presetRangeForClear.from;
                        //         this.selectedMaxDate = presetRangeForClear.to;
                        //         this.activePreset = selectedPreset;
                        //         this.lastPresetSetting = selectedPreset;
                        //         this.hasManualSelection = false;
                                
                        //         // Apply the filter
                        //         this.isUserInitiatedChange = true;
                        //         this.applyDateFilter(category.source, presetRangeForClear.from, presetRangeForClear.to);
                        //         this.lastAppliedFilterHash = this.createFilterHash([{
                        //             $schema: "http://powerbi.com/product/schema#advanced",
                        //             target: { 
                        //                 table: this.extractTableAndColumn(category.source).tableName, 
                        //                 column: this.extractTableAndColumn(category.source).columnName 
                        //             },
                        //             logicalOperator: "And",
                        //             conditions: [
                        //                 { operator: "GreaterThanOrEqual", value: presetRangeForClear.from.toISOString() },
                        //                 { operator: "LessThanOrEqual", value: presetRangeForClear.to.toISOString() }
                        //             ]
                        //         } as powerbi.IFilter], category.source);
                                
                        //         this.needsReRender = true;
                                
                        //         setTimeout(() => {
                        //             this.isUserInitiatedChange = false;
                        //         }, 100);
                        //     } else {
                        //         // Honor external filter if it matches preset or no preset is set
                        //         console.log("🔄 SYNC UPDATE: Honoring external filter");
                        //         this.selectedMinDate = externalFilterRange.minDate;
                        //         this.selectedMaxDate = externalFilterRange.maxDate;
                        //         this.hasManualSelection = true;
                        //         this.shouldApplyFilter = false;
                        //         this.needsReRender = true;
                        //         this.lastAppliedFilterHash = externalFilterHash;

                        //         if (presetRangeForClear) {
                        //             const matchesPreset = Math.abs(externalFilterRange.minDate.getTime() - presetRangeForClear.from.getTime()) <= 1000 &&
                        //                 Math.abs(externalFilterRange.maxDate.getTime() - presetRangeForClear.to.getTime()) <= 1000;
                        //             if (matchesPreset && selectedPreset && selectedPreset !== "none") {
                        //                 this.activePreset = selectedPreset;
                        //                 this.lastPresetSetting = selectedPreset;
                        //             }
                        //         }
                        //     }
                        // }

                        // Sync update 3
                        else if (externalFilterRange && this.isSyncUpdateFlag && !this.isUserInitiatedChange && !this.isRestoringBookmark) {
                                console.log("🔄 SYNC UPDATE: Checking source of filter");
                                
                                // Check if external filter matches what USER selected through this visual
                                const matchesUserSelection = this.lastUserSelectedMin && this.lastUserSelectedMax &&
                                    Math.abs(externalFilterRange.minDate.getTime() - this.lastUserSelectedMin.getTime()) <= 1000 &&
                                    Math.abs(externalFilterRange.maxDate.getTime() - this.lastUserSelectedMax.getTime()) <= 1000;
                                
                                // Check if external filter matches live preset
                                const matchesLivePreset = presetRangeForClear && 
                                    Math.abs(externalFilterRange.minDate.getTime() - presetRangeForClear.from.getTime()) <= 1000 &&
                                    Math.abs(externalFilterRange.maxDate.getTime() - presetRangeForClear.to.getTime()) <= 1000;
                                
                                // If it matches user selection → Honor it (page navigation)
                                if (matchesUserSelection) {
                                    console.log("✅ Matches user selection - preserving across page change");
                                    this.selectedMinDate = externalFilterRange.minDate;
                                    this.selectedMaxDate = externalFilterRange.maxDate;
                                    this.hasManualSelection = true;
                                    this.shouldApplyFilter = false;
                                    this.needsReRender = true;
                                    this.lastAppliedFilterHash = externalFilterHash;
                                }
                                // If it matches live preset → Honor it
                                else if (matchesLivePreset) {
                                    console.log("✅ Matches live preset - honoring");
                                    this.selectedMinDate = externalFilterRange.minDate;
                                    this.selectedMaxDate = externalFilterRange.maxDate;
                                    this.activePreset = selectedPreset;
                                    this.lastPresetSetting = selectedPreset;
                                    this.needsReRender = true;
                                    this.lastAppliedFilterHash = externalFilterHash;
                                }
                                // If it matches NEITHER → Clear All was clicked → Apply live preset
                                else if (presetRangeForClear && selectedPreset && selectedPreset !== "none") {
                                    console.log("🔘 CLEAR ALL DETECTED: Applying live preset");
                                    
                                    this.selectedMinDate = presetRangeForClear.from;
                                    this.selectedMaxDate = presetRangeForClear.to;
                                    this.activePreset = selectedPreset;
                                    this.lastPresetSetting = selectedPreset;
                                    this.hasManualSelection = false;
                                    
                                    // Reset user selection tracking
                                    this.lastUserSelectedMin = null;
                                    this.lastUserSelectedMax = null;
                                    
                                    this.isUserInitiatedChange = true;
                                    this.applyDateFilter(category.source, presetRangeForClear.from, presetRangeForClear.to);
                                    this.lastAppliedFilterHash = this.createFilterHash([{
                                        $schema: "http://powerbi.com/product/schema#advanced",
                                        target: { 
                                            table: this.extractTableAndColumn(category.source).tableName, 
                                            column: this.extractTableAndColumn(category.source).columnName 
                                        },
                                        logicalOperator: "And",
                                        conditions: [
                                            { operator: "GreaterThanOrEqual", value: presetRangeForClear.from.toISOString() },
                                            { operator: "LessThanOrEqual", value: presetRangeForClear.to.toISOString() }
                                        ]
                                    } as powerbi.IFilter], category.source);
                                    
                                    this.needsReRender = true;
                                    
                                    setTimeout(() => {
                                        this.isUserInitiatedChange = false;
                                    }, 100);
                                }
                                // No preset configured → honor external filter
                                else {
                                    console.log("🔄 No preset - honoring external filter");
                                    this.selectedMinDate = externalFilterRange.minDate;
                                    this.selectedMaxDate = externalFilterRange.maxDate;
                                    this.hasManualSelection = true;
                                    this.needsReRender = true;
                                    this.lastAppliedFilterHash = externalFilterHash;
                                }
                            }


                        // External filter exists
                        else if (externalFilterRange && !this.isUserInitiatedChange && !this.isRestoringBookmark && !detectedBookmarkRestore && !wasInitialRender) {
                            const selectionChanged =
                                !this.selectedMinDate ||
                                !this.selectedMaxDate ||
                                Math.abs(this.selectedMinDate.getTime() - externalFilterRange.minDate.getTime()) > 1000 ||
                                Math.abs(this.selectedMaxDate.getTime() - externalFilterRange.maxDate.getTime()) > 1000;

                            if (selectionChanged) {
                                console.log("📋 Honoring external filter");

                                this.selectedMinDate = externalFilterRange.minDate;
                                this.selectedMaxDate = externalFilterRange.maxDate;
                                this.hasManualSelection = true;
                                this.lastAppliedFilterHash = externalFilterHash;
                                this.shouldApplyFilter = false;
                                this.needsReRender = true;

                                if (presetRangeForClear) {
                                    const matchesPreset = Math.abs(externalFilterRange.minDate.getTime() - presetRangeForClear.from.getTime()) <= 1000 &&
                                        Math.abs(externalFilterRange.maxDate.getTime() - presetRangeForClear.to.getTime()) <= 1000;
                                    if (matchesPreset && selectedPreset && selectedPreset !== "none") {
                                        this.activePreset = selectedPreset;
                                        this.lastPresetSetting = selectedPreset;
                                    }
                                }
                            }
                        }
                        
                        
                        // No external filter
                        else if (!externalFilterRange && !this.isUserInitiatedChange) {
                            if (wasInitialRender && presetRangeForClear && selectedPreset && selectedPreset !== "none") {
                                this.selectedMinDate = presetRangeForClear.from;
                                this.selectedMaxDate = presetRangeForClear.to;
                                this.activePreset = selectedPreset;
                                this.lastPresetSetting = selectedPreset;
                                this.hasManualSelection = false;
                                this.shouldApplyFilter = true;
                                this.applyDateFilter(category.source, presetRangeForClear.from, presetRangeForClear.to);
                                this.lastAppliedFilterHash = this.createFilterHash([{
                                    $schema: "http://powerbi.com/product/schema#advanced",
                                    target: { 
                                        table: this.extractTableAndColumn(category.source).tableName, 
                                        column: this.extractTableAndColumn(category.source).columnName 
                                    },
                                    logicalOperator: "And",
                                    conditions: [
                                        { operator: "GreaterThanOrEqual", value: presetRangeForClear.from.toISOString() },
                                        { operator: "LessThanOrEqual", value: presetRangeForClear.to.toISOString() }
                                    ]
                                } as powerbi.IFilter], category.source);
                                this.shouldApplyFilter = false;
                                this.needsReRender = true;
                            } else if (!this.selectedMinDate || !this.selectedMaxDate) {
                                this.selectedMinDate = minDate;
                                this.selectedMaxDate = maxDate;
                                this.needsReRender = true;
                            }
                        }

                        console.log("selectedMinDate", this.selectedMinDate);
                        console.log("selectedMaxDate", this.selectedMaxDate);
                        console.log("presetRangeForClear",presetRangeForClear);
                        console.log("externalFilterRange",externalFilterRange);
                        console.log("externalFilterHash",externalFilterHash);

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

                        if (this.needsReRender && this.reactSliderWrapper) {
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

    // ============================================================================
    // 🔑 SYNC FILTER HELPERS
    // ============================================================================
    // Helper to create a hash of filter state for comparison
    private createFilterHash(filters: powerbi.IFilter[], source?: powerbi.DataViewMetadataColumn): string | null {
        if (!filters || !filters.length || !source) {
            return null;
        }

        const { tableName, columnName } = this.extractTableAndColumn(source);
        const relevantFilters = filters
            .filter((f: any) => {
                const target = f?.target;
                return target && target.table === tableName && target.column === columnName;
            })
            .map((f: any) => {
                const conditions = f?.conditions || [];
                return {
                    logicalOperator: f?.logicalOperator,
                    conditions: conditions.map((c: any) => ({
                        operator: c?.operator,
                        value: c?.value
                    }))
                };
            });

        return relevantFilters.length > 0 ? JSON.stringify(relevantFilters) : null;
    }

    // Helper to check if an external filter represents a sync update
    private checkIsSyncUpdate(externalFilterHash: string | null): boolean {
        // It's a sync update if:
        // 1. External filter exists (not null)
        // 2. We have a last applied filter hash
        // 3. External filter hash differs from what we last applied
        // 4. This is NOT a user-initiated change
        return !!externalFilterHash &&
            !!this.lastAppliedFilterHash &&
            externalFilterHash !== this.lastAppliedFilterHash &&
            !this.isUserInitiatedChange;
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


        this.selectedMinDate = normalizedMin;
        this.selectedMaxDate = normalizedMax;

        // Mark this as a user-initiated change - filter MUST be applied
        this.isUserInitiatedChange = true;
        this.shouldApplyFilter = true;

        // If this is a manual selection (not from preset change), mark it
        if (!isPresetChange) {
            this.hasManualSelection = true;
            // console.log("Manual selection detected - preset will not override this");
        }

        // Use provided source or current data source
        const dataSource = source || this.currentDataSource;
        if (dataSource) {
            // Always apply filter for user-initiated changes
            this.applyDateFilter(dataSource, normalizedMin, normalizedMax);
            
            // Update last applied filter hash to prevent sync loops
            // This ensures that when our filter propagates to other pages, they don't treat it as a sync update
            const appliedFilter: powerbi.IFilter = {
                $schema: "http://powerbi.com/product/schema#advanced",
                target: { 
                    table: this.extractTableAndColumn(dataSource).tableName, 
                    column: this.extractTableAndColumn(dataSource).columnName 
                },
                logicalOperator: "And",
                conditions: [
                    { operator: "GreaterThanOrEqual", value: normalizedMin.toISOString() },
                    { operator: "LessThanOrEqual", value: normalizedMax.toISOString() }
                ]
            };
            this.lastAppliedFilterHash = this.createFilterHash([appliedFilter], dataSource);
        } else {
            // console.warn("No data source available when trying to apply filter");
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
            // Use merge to combine with other slicers' filters
            // This ensures the filter is applied to other visuals while respecting other filters
            this.host.applyJsonFilter(
                advancedFilter,
                "general",
                "filter",
                powerbi.FilterAction.merge
            );
        } else {
            // console.log("Filter application skipped - shouldApplyFilter is false");
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

            // console.log("currentMaxToUse", currentMaxToUse);
            // console.log("currentMinToUse", currentMinToUse);


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
                    // Mark as manual selection and user-initiated change immediately
                    // This prevents post-restore correction from overriding manual selections
                    this.lastUserSelectedMin = minDate;
                    this.lastUserSelectedMax = maxDate;
                    this.hasManualSelection = true;
                    this.isUserInitiatedChange = true;
                    this.selectedMinDate = newMinDate;
                    this.selectedMaxDate = newMaxDate;

                    if (this.currentDataSource) {
                        this.shouldApplyFilter = true;
                        this.applyDateFilter(this.currentDataSource, newMinDate, newMaxDate);
                    }

                    // Reset isUserInitiatedChange after a delay to allow update cycle to complete
                    setTimeout(() => {
                        this.isUserInitiatedChange = false;
                    }, 200);
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

