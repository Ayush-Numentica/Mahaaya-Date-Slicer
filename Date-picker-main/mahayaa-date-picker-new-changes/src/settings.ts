"use strict";

import { formattingSettings } from "powerbi-visuals-utils-formattingmodel";
import { dataViewObjectsParser } from "powerbi-visuals-utils-dataviewutils";

import FormattingSettingsCard = formattingSettings.SimpleCard;
import FormattingSettingsSlice = formattingSettings.Slice;
import FormattingSettingsModel = formattingSettings.Model;

/**
 * Data Point Formatting Card (Color options)
 */
class DataPointCardSettings extends FormattingSettingsCard {
    cardColor = new formattingSettings.ColorPicker({
        name: "cardColor",
        displayName: "Card Color",
        value: { value: "" }
    });

    showAllDataPoints = new formattingSettings.ToggleSwitch({
        name: "showAllDataPoints",
        displayName: "Show all",
        value: true
    });

    dateBoxColor = new formattingSettings.ColorPicker({
        name: "dateBoxColor",
        displayName: "Date Box color",
        value: { value: "" }
    });

    fontColor = new formattingSettings.ColorPicker({
        name: "fontColor",
        displayName: "Font Color",
        value: { value: "" }
    });

    fontSize = new formattingSettings.NumUpDown({
        name: "fontSize",
        displayName: "Text Size",
        value: 12
    });

    name: string = "dataPoint";
    displayName: string = "Data colors";
    slices: Array<FormattingSettingsSlice> = [
        this.cardColor,
        this.showAllDataPoints,
        this.dateBoxColor,
        this.fontColor,
        this.fontSize
    ];
}

/**
 * Presets Card (Date filter presets)
 */
class PresetsCardSettings extends FormattingSettingsCard {
    // Single dropdown instead of multiple toggles
    toShowHeader = new formattingSettings.ToggleSwitch({
        name: "toShowHeader",
        displayName: "Show the Header",
        value: true
    });
    preset = new formattingSettings.AutoDropdown({
        name: "preset",            // must match capabilities.json
        displayName: "Chart Filter",
        value: "none"              // default
    });
    selectionStyle = new formattingSettings.AutoDropdown({
        name: "Selection Style",            // must match capabilities.json
        displayName: "Selection Style",
        value: "slider"              // default
    });
    toggleOption = new formattingSettings.ToggleSwitch({
        name: "toggleOption",
        displayName: "Pop Up Mode",
        value: false
    });
    recalculatePresetOnBookmark = new formattingSettings.ToggleSwitch({
        name: "recalculatePresetOnBookmark",
        displayName: "Recalculate preset on bookmark restore",
        value: true
    });

    name: string = "presets";     // must match capabilities.json
    displayName: string = "Presets";
    slices: Array<FormattingSettingsSlice> = [this.toShowHeader,this.preset, this.selectionStyle,this.toggleOption, this.recalculatePresetOnBookmark];
}



/**
 * Visual Formatting Settings Model (new API for formatting pane)
 */
export class VisualFormattingSettingsModel extends FormattingSettingsModel {
    dataPointCard = new DataPointCardSettings();
    presetsCard = new PresetsCardSettings();

    cards = [this.dataPointCard, this.presetsCard];
}

/**
 * Legacy parser (so you can read settings in visual.ts)
 */
// export class PresetsSettings {
//     public showThisMonth: boolean = false;
//     public showLastMonth: boolean = false;
//     public showLast7Days: boolean = false;
//     public showLast30Days: boolean = false;
// }

// export class VisualSettings extends dataViewObjectsParser.DataViewObjectsParser {
//     public presets: PresetsSettings = new PresetsSettings();
// }
