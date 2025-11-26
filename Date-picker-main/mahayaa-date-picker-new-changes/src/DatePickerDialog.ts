import powerbi from "powerbi-visuals-api";
import DialogConstructorOptions = powerbi.extensibility.visual.DialogConstructorOptions;
import DialogAction = powerbi.DialogAction;
// React imports
import { createRoot } from 'react-dom/client';
import * as React from 'react';
import { format, parse, isValid } from "date-fns";
import { Calendar as CalendarIcon } from "lucide-react";

// Import the Calendar component
import { Calendar } from "../components/ui/calendar";

export class DatePickerDialog {
    static id = "DatePickerDialog";

    constructor(options: DialogConstructorOptions, initialState: object) {
        const host = options.host;
        let pickedFromDate: Date | undefined;
        let pickedToDate: Date | undefined;

        const startFromDate = new Date(initialState['fromDate']);
        const startToDate = new Date(initialState['toDate']);
        const minDate = new Date(initialState['minDate']);
        const maxDate = new Date(initialState['maxDate']);
        const backgroundColor = (initialState as any)['backgroundColor'] || '#FFFFFF';

        // Helper utilities for month math
        const startOfMonth = (date: Date) => new Date(date.getFullYear(), date.getMonth(), 1);
        const addMonths = (date: Date, months: number) => {
            const result = new Date(date);
            result.setMonth(result.getMonth() + months);
            return startOfMonth(result);
        };
        const compareMonths = (a: Date, b: Date) => startOfMonth(a).getTime() - startOfMonth(b).getTime();

        const earliestStartMonth = startOfMonth(minDate);
        const getLatestStartMonth = () => {
            const raw = addMonths(maxDate, -2);
            return compareMonths(raw, earliestStartMonth) < 0 ? earliestStartMonth : raw;
        };

        const clampStartMonth = (candidate: Date) => {
            const latestStartMonth = getLatestStartMonth();
            if (compareMonths(candidate, earliestStartMonth) < 0) {
                return startOfMonth(earliestStartMonth);
            }
            if (compareMonths(candidate, latestStartMonth) > 0) {
                return startOfMonth(latestStartMonth);
            }
            return startOfMonth(candidate);
        };

        const getCenteredStartMonth = (fromValue?: Date, toValue?: Date) => {
            const anchor = fromValue || toValue || startFromDate || minDate;
            const centerMonth = startOfMonth(anchor);
            // return clampStartMonth(addMonths(centerMonth, -1));
            return addMonths(centerMonth, -1);
        };

        // Initialize with provided dates
        pickedFromDate = startFromDate;
        pickedToDate = startToDate;
        let currentCalendarMonth = getCenteredStartMonth(pickedFromDate, pickedToDate);

        // State for text inputs - using regular variables instead of React state
        let fromText = startFromDate ? format(startFromDate, "MM/dd/yyyy") : "";
        let toText = startToDate ? format(startToDate, "MM/dd/yyyy") : "";

        // Re-render function
        const reRender = () => {
            root.render(createDialogContent());
        };

        const shiftCalendarMonth = (delta: number) => {
            // const nextMonth = clampStartMonth(addMonths(currentCalendarMonth, delta));
            // if (nextMonth.getTime() !== currentCalendarMonth.getTime()) {
            //     currentCalendarMonth = nextMonth;
            //     reRender();
            // }
            const nextMonth = addMonths(currentCalendarMonth, delta);
            currentCalendarMonth = nextMonth;
            reRender();
        };

        const applySelection = (fromValue: Date, toValue: Date, closeDialog: boolean = false) => {
            const normalizedFrom = new Date(fromValue);
            normalizedFrom.setHours(0, 0, 0, 0);

            const normalizedTo = new Date(toValue);
            normalizedTo.setHours(23, 59, 59, 999);

            pickedFromDate = normalizedFrom;
            pickedToDate = normalizedTo;

            fromText = format(normalizedFrom, "MM/dd/yyyy");
            toText = format(normalizedTo, "MM/dd/yyyy");

            // currentCalendarMonth = getCenteredStartMonth(normalizedFrom, normalizedTo);
            // Center calendar on selection without clamping
            const anchor = normalizedFrom || normalizedTo || startFromDate || minDate;
            currentCalendarMonth = addMonths(startOfMonth(anchor), -1);

            host.setResult({ fromDate: normalizedFrom, toDate: normalizedTo });

            if (closeDialog) {
                host.close(DialogAction.OK, { fromDate: normalizedFrom, toDate: normalizedTo });
            } else {
                reRender();
            }
        };

        // Update temp dates when text inputs change
        const handleFromTextChange = (value: string) => {
            fromText = value;

            // ⛔ Don't parse until full date format is entered
            if (value.length < 10) {
                reRender();
                return;
            }

            const parsed = parse(value, "MM/dd/yyyy", new Date());

            if (isValid(parsed)) {
                const normalizedTo = pickedToDate || parsed;
                applySelection(parsed, normalizedTo);
            } else {
                reRender();
            }
        };


        const handleToTextChange = (value: string) => {
            toText = value;

            // ⛔ Wait until user types full date (MM/dd/yyyy)
            if (value.length < 10) {
                reRender();
                return;
            }

            const parsed = parse(value, "MM/dd/yyyy", new Date());

            if (isValid(parsed)) {
                const normalizedFrom = pickedFromDate || parsed;
                applySelection(normalizedFrom, parsed);
            } else {
                reRender();
            }
        };


        const handleCalendarSelect = (range: { from?: Date; to?: Date } | undefined) => {
            if (!range) return;
            const normalizedFrom = range.from || range.to;
            const normalizedTo = range.to || range.from;

            if (normalizedFrom && normalizedTo) {
                applySelection(normalizedFrom, normalizedTo);
            } else {
                reRender();
            }
        };

        // Calculate responsive sizes based on viewport
        const getResponsiveSizes = () => {
            if (typeof window === 'undefined') {
                return {
                    baseFont: 14,
                    titleFont: 20,
                    padding: 14,
                    gap: 12,
                    buttonPadding: '8px 16px',
                    inputPadding: '8px 12px',
                    borderRadius: 6,
                    margin: 16
                };
            }

            const viewportWidth = window.innerWidth;
            const viewportHeight = window.innerHeight;
            const viewportArea = viewportWidth * viewportHeight;

            // Base sizes for 1920x1080 (standard desktop)
            const baseArea = 1920 * 1080;
            const scaleFactor = Math.sqrt(viewportArea / baseArea);

            // Clamp scale factor between 0.7 and 1.3 for reasonable bounds
            const clampedScale = Math.max(0.7, Math.min(1.3, scaleFactor));

            return {
                baseFont: Math.round(14 * clampedScale),
                titleFont: Math.round(20 * clampedScale),
                padding: Math.round(14 * clampedScale),
                gap: Math.round(12 * clampedScale),
                buttonPadding: `${Math.round(8 * clampedScale)}px ${Math.round(16 * clampedScale)}px`,
                inputPadding: `${Math.round(8 * clampedScale)}px ${Math.round(12 * clampedScale)}px`,
                borderRadius: Math.round(6 * clampedScale),
                margin: Math.round(16 * clampedScale),
                labelMargin: Math.round(8 * clampedScale)
            };
        };

        // Create dialog content function
        const createDialogContent = () => {
            const sizes = getResponsiveSizes();
            const datasetMinDate = minDate;
            const datasetMaxDate = maxDate;
            const navButtonStyle: React.CSSProperties = {
                border: '1px solid #d0d0d0',
                borderRadius: `${sizes.borderRadius}px`,
                background: '#ffffff',
                width: `${Math.round(40 * (sizes.baseFont / 14))}px`,
                height: `${Math.round(40 * (sizes.baseFont / 14))}px`,
                fontSize: `${sizes.baseFont * 1.4}px`,
                fontWeight: 600,
                cursor: 'pointer',
                color: '#0078d4',
                display: 'flex',
                alignItems: 'center',
                justifyContent: 'center',
                boxShadow: '0 1px 2px rgba(0,0,0,0.08)'
            };
            // const latestStartMonth = getLatestStartMonth();
            // const canGoPrev = compareMonths(currentCalendarMonth, earliestStartMonth) > 0;
            // const canGoNext = compareMonths(currentCalendarMonth, latestStartMonth) < 0;

            return React.createElement(
                'div',
                {
                    style: {
                        padding: `${sizes.padding}px`,
                        fontFamily: 'Segoe UI, sans-serif',
                        width: '100%',
                        maxWidth: '800px',
                        // height: '100%',
                        // maxHeight: '100vh',
                        overflowY: 'hidden',
                        background: backgroundColor,
                        display: 'flex',
                        flexDirection: 'column',
                        boxSizing: 'border-box',
                        // border:"1px solid black",
                        position:"fixed"
                    }
                },
                [
                    // Responsive styles for calendar
                    React.createElement('style', {
                        key: 'responsive-calendar-styles',
                        dangerouslySetInnerHTML: {
                            __html: `
                                .responsive-calendar {
                                    font-size: ${sizes.baseFont * 0.9}px !important;
                                }
                                .responsive-calendar .rdp-caption_label {
                                    font-size: ${sizes.baseFont}px !important;
                                }
                                .responsive-calendar .rdp-head_cell {
                                    font-size: ${sizes.baseFont * 0.8}px !important;
                                }
                                .responsive-calendar .rdp-button {
                                    font-size: ${sizes.baseFont * 0.9}px !important;
                                }
                                .responsive-calendar .rdp-day {
                                    font-size: ${sizes.baseFont * 0.9}px !important;
                                }
                            `
                        }
                    }),
                    // Header
                    React.createElement('div', {
                        key: 'header',
                        style: { marginBottom: `${sizes.margin}px`, flexShrink: 0 }
                    }, [
                        React.createElement('h2', {
                            key: 'title',
                            style: {
                                fontSize: `${sizes.titleFont}px`,
                                fontWeight: '600',
                                color: '#333',
                                margin: `0 0 ${sizes.labelMargin}px 0`
                            }
                        }, 'Select Date Range')
                    ]),


                     // Preset buttons
                    React.createElement('div', {
                        key: 'presets',
                        style: {
                            display: 'flex',
                            justifyContent: 'space-between',
                            gap: `${sizes.gap}px`,
                            marginBottom: `${sizes.margin}px`,
                            flexWrap: 'wrap',
                            flexShrink: 0
                        }
                    }, [
                        React.createElement('button', {
                            key: 'today',
                            onClick: () => {
                                const today = new Date();
                                applySelection(today, today, false);
                            },
                            style: {
                                flex: 1,
                                padding: sizes.buttonPadding,
                                border: '1px solid #0078d4',
                                borderRadius: `${sizes.borderRadius}px`,
                                background: 'white',
                                color: '#0078d4',
                                fontSize: `${sizes.baseFont}px`,
                                fontWeight: '500',
                                cursor: 'pointer',
                                transition: 'all 0.2s ease'
                            }
                        }, 'Today'),
                        React.createElement('button', {
                            key: 'yesterday',
                            onClick: () => {
                                const yesterday = new Date();
                                yesterday.setDate(yesterday.getDate() - 1);
                                applySelection(yesterday, yesterday, false);
                            },
                            style: {
                                flex: 1,
                                padding: sizes.buttonPadding,
                                border: '1px solid #0078d4',
                                borderRadius: `${sizes.borderRadius}px`,
                                background: 'white',
                                color: '#0078d4',
                                fontSize: `${sizes.baseFont}px`,
                                fontWeight: '500',
                                cursor: 'pointer',
                                transition: 'all 0.2s ease'
                            }
                        }, 'Yesterday'),
                        React.createElement('button', {
                            key: 'min-date',
                            onClick: () => {
                                applySelection(minDate, minDate, false);
                                // const single = new Date(datasetMinDate);
                                // single.setHours(0, 0, 0, 0);
                                // pickedFromDate = single;
                                // pickedToDate = undefined;
                                // fromText = format(single, "MM/dd/yyyy");
                                // toText = "";
                                // host.setResult({ fromDate: single, toDate: single });
                                // reRender();
                            },
                            style: {
                                flex: 1,
                                padding: sizes.buttonPadding,
                                border: '1px solid #0078d4',
                                borderRadius: `${sizes.borderRadius}px`,
                                background: 'white',
                                color: '#0078d4',
                                fontSize: `${sizes.baseFont}px`,
                                fontWeight: '500',
                                cursor: 'pointer',
                                transition: 'all 0.2s ease'
                            }
                        }, 'Min Date'),
                        React.createElement('button', {
                            key: 'max-date',
                            onClick: () => {
                                applySelection(maxDate, maxDate, false);
                                // const single = new Date(datasetMaxDate);
                                // single.setHours(0, 0, 0, 0);
                                // pickedFromDate = single;
                                // pickedToDate = undefined;
                                // fromText = format(single, "MM/dd/yyyy");
                                // toText = "";
                                // host.setResult({ fromDate: single, toDate: single });
                                // reRender();
                            },
                            style: {
                                flex: 1,
                                padding: sizes.buttonPadding,
                                border: '1px solid #0078d4',
                                borderRadius: `${sizes.borderRadius}px`,
                                background: 'white',
                                color: '#0078d4',
                                fontSize: `${sizes.baseFont}px`,
                                fontWeight: '500',
                                cursor: 'pointer',
                                transition: 'all 0.2s ease'
                            }
                        }, 'Max Date')
                    ]),

                    // Text inputs
                    React.createElement('div', {
                        key: 'inputs',
                        style: {
                            marginBottom: `${sizes.margin}px`,
                            flexShrink: 0,
                            width: '100%'
                        }
                    }, [
                        React.createElement('div', {
                            key: 'input-row',
                            style: {
                                display: 'flex',
                                alignItems: 'flex-end',
                                justifyContent: 'center',
                                gap: `${sizes.gap * 1.5}px`,
                                maxWidth: '620px',
                                margin: '0 auto'
                            }
                        }, [
                            React.createElement('button', {
                                key: 'nav-prev',
                                onClick: () => shiftCalendarMonth(-1),
                                // disabled: !canGoPrev,
                                // style: {
                                //     ...navButtonStyle,
                                //     opacity: canGoPrev ? 1 : 0.35,
                                //     cursor: canGoPrev ? 'pointer' : 'not-allowed'
                                // },
                                style: navButtonStyle,
                                'aria-label': 'Previous month'
                            }, '‹'),
                            React.createElement('div', {
                                key: 'from-input',
                                style: {
                                    display: 'flex',
                                    flexDirection: 'column',
                                    alignItems: 'flex-start',
                                    flex: '1'
                                }
                            }, [
                                React.createElement('label', {
                                    key: 'from-label',
                                    style: {
                                        display: 'block',
                                        fontSize: `${sizes.baseFont}px`,
                                        fontWeight: '500',
                                        color: '#333',
                                        marginBottom: `${sizes.labelMargin}px`
                                    }
                                }, 'From Date'),
                                React.createElement('input', {
                                    key: 'from-field',
                                    type: 'text',
                                    value: fromText,
                                    onChange: (e: any) => handleFromTextChange(e.target.value),
                                    placeholder: 'MM/DD/YYYY',
                                    style: {
                                        width: '100%',
                                        minWidth: `${Math.round(140 * (sizes.baseFont / 14))}px`,
                                        padding: sizes.inputPadding,
                                        border: '1px solid #d0d0d0',
                                        borderRadius: `${sizes.borderRadius}px`,
                                        fontSize: `${sizes.baseFont}px`,
                                        fontFamily: 'Segoe UI, sans-serif',
                                        boxSizing: 'border-box'
                                    }
                                })
                            ]),
                            React.createElement('div', {
                                key: 'to-input',
                                style: {
                                    display: 'flex',
                                    flexDirection: 'column',
                                    alignItems: 'flex-start',
                                    flex: '1'
                                }
                            }, [
                                React.createElement('label', {
                                    key: 'to-label',
                                    style: {
                                        display: 'block',
                                        fontSize: `${sizes.baseFont}px`,
                                        fontWeight: '500',
                                        color: '#333',
                                        marginBottom: `${sizes.labelMargin}px`
                                    }
                                }, 'To Date'),
                                React.createElement('input', {
                                    key: 'to-field',
                                    type: 'text',
                                    value: toText,
                                    onChange: (e: any) => handleToTextChange(e.target.value),
                                    placeholder: 'MM/DD/YYYY',
                                    style: {
                                        width: '100%',
                                        minWidth: `${Math.round(140 * (sizes.baseFont / 14))}px`,
                                        padding: sizes.inputPadding,
                                        border: '1px solid #d0d0d0',
                                        borderRadius: `${sizes.borderRadius}px`,
                                        fontSize: `${sizes.baseFont}px`,
                                        fontFamily: 'Segoe UI, sans-serif',
                                        boxSizing: 'border-box'
                                    }
                                })
                            ]),
                            React.createElement('button', {
                                key: 'nav-next',
                                onClick: () => shiftCalendarMonth(1),
                                // disabled: !canGoNext,
                                // style: {
                                //     ...navButtonStyle,
                                //     opacity: canGoNext ? 1 : 0.35,
                                //     cursor: canGoNext ? 'pointer' : 'not-allowed'
                                // },
                                style: navButtonStyle,
                                'aria-label': 'Next month'
                            }, '›')
                        ])
                    ]),

                   

                    // Calendar with responsive styling
                    (() => {
                        // Calculate available height for calendar
                        // Dialog height (400px) - padding (top+bottom) - header - inputs - buttons - margins
                        const dialogHeight = 350;
                        const buttonPaddingTop = parseInt(sizes.buttonPadding.split(' ')[0].replace('px', '')) || 8;
                        const usedHeight = sizes.padding * 2 + // top and bottom padding
                            sizes.titleFont + sizes.labelMargin + // header
                            (sizes.baseFont + sizes.labelMargin) * 2 + sizes.margin * 2 + // inputs section
                            (sizes.baseFont + buttonPaddingTop * 2) + sizes.margin + // buttons section
                            sizes.labelMargin; // calendar margin
                        const availableHeight = dialogHeight - usedHeight;

                        // Estimate calendar height: 2 months side by side
                        // Each month: caption (~40px) + 6 weeks * cell-size + gaps
                        const baseCellSize = 40;
                        const estimatedCalendarHeight = (sizes.titleFont + 20) + (6 * baseCellSize) + 20; // caption + weeks + gaps
                        const estimatedCalendarWidth = (baseCellSize * 7) * 3 + 60; // 3 months side by side

                        // Calculate scale factor to fit in available space
                        const heightScale = availableHeight / estimatedCalendarHeight;
                        const widthScale = (800 - sizes.padding * 2) / estimatedCalendarWidth;
                        const calendarScale = Math.min(heightScale, widthScale, 1); // Don't scale up, only down

                        const scaledCellSize = Math.round(baseCellSize * calendarScale * (sizes.baseFont / 14));
                        const scaledFontSize = sizes.baseFont * 1.1 * calendarScale;

                        return React.createElement('div', {
                            key: 'calendar',
                            style: {
                                display: 'flex',
                                justifyContent: 'center',
                                alignItems: 'flex-start',
                                flex: '1 1 auto',
                                minHeight: 0,
                                overflow: 'hidden',
                                marginTop: `${sizes.labelMargin}px`,
                                // Responsive calendar sizing using CSS variables
                                '--cell-size': `${scaledCellSize}px`,
                                fontSize: `${scaledFontSize}px`
                            } as React.CSSProperties
                        }, React.createElement('div', {
                            style: {
                                transform: `scale(${calendarScale})`,
                                transformOrigin: 'top center',
                                width: `${100 / calendarScale}%`,
                                height: `${100 / calendarScale}%`
                            }
                        }, React.createElement(Calendar, {
                            mode: "range",
                            defaultMonth: currentCalendarMonth,
                            month: currentCalendarMonth,
                            onMonthChange: (newMonth: Date) => {
                                // currentCalendarMonth = clampStartMonth(startOfMonth(newMonth));
                                currentCalendarMonth = startOfMonth(newMonth);
                                reRender();
                            },
                            selected: { from: pickedFromDate, to: pickedToDate },
                            onSelect: handleCalendarSelect,
                            numberOfMonths: 3,
                            initialFocus: true,
                            className: 'responsive-calendar',
                            classNames: { nav: 'hidden' }
                        }),
                            // React.createElement(Calendar, {
                            //     mode: "range",
                            //     defaultMonth: pickedFromDate || startFromDate,
                            //     selected: { from: pickedFromDate, to: pickedToDate },
                            //     onSelect: handleCalendarSelect,
                            //     numberOfMonths: 2,
                            //     initialFocus: true,
                            //     className: "responsive-calendar"
                            // }),

                            // ⬇⬇ ADD THIS BLOCK EXACTLY AFTER THE CALENDAR ⬇⬇
                            React.createElement("style", {
                                dangerouslySetInnerHTML: {
                                    __html: `
                                                .responsive-calendar { 
                                                    font-size: 15px !important;       /* Bigger overall text */
                                                }
                                                .responsive-calendar .rdp-caption_label { 
                                                    font-size: 17px !important;       /* Month + Year bigger */
                                                    font-weight: 600;
                                                }
                                                .responsive-calendar .rdp-button { 
                                                    font-size: 22px !important;       /* Month + Year bigger */
                                                    font-weight: 600;
                                                    display: none !important
                                                }
                                                .responsive-calendar .rdp-day, 
                                                .responsive-calendar .rdp-weekday {
                                                    font-size: 15px !important;       /* Individual dates bigger */
                                                }
                                            `
                                }
                            })

                        ));
                    })()

                ]
            );
        };

        // Dialog rendering implementation
        const root = createRoot(options.element);
        root.render(createDialogContent());

        // Handle Enter key to close dialog with current selection
        document.addEventListener('keydown', e => {
            if (e.code == 'Enter' && pickedFromDate && pickedToDate) {
                host.close(DialogAction.OK, { fromDate: pickedFromDate, toDate: pickedToDate });
            }
        });
    }
}

export class DatePickerDialogResult {
    fromDate: Date;
    toDate: Date;
}

// Register the dialog
globalThis.dialogRegistry = globalThis.dialogRegistry || {};
globalThis.dialogRegistry[DatePickerDialog.id] = DatePickerDialog;
