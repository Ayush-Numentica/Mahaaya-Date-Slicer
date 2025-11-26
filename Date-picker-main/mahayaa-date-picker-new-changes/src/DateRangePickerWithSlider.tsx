import React, { useState, useCallback, useEffect, memo, useRef } from "react";
import { DateRangePicker } from "./DateRangePicker";
import { Calendar } from "../components/ui/calendar";
import { DatePickerDialogResult } from "./DatePickerDialog";
import { toDate } from "date-fns";
import type { DateRange } from 'react-day-picker';
import { format, parse, isValid } from "date-fns";





interface DateRangeSliderProps {
  minDate: Date;
  maxDate: Date;
  currentMinDate: Date;
  currentMaxDate: Date;
  onDateChange: (minDate: Date, maxDate: Date) => void;
  formatDateForSlider: (date: Date) => string;
  onOpenDialog: (
    fromDate: Date,
    toDate: Date,
    minDate: Date,
    maxDate: Date
  ) => void;
  // When true, hide the slider UI and expose only the popup picker button
  usePopupOnly?: boolean;
  datePickerType?: string;
  // Styling passed from formatting pane
  inputFontSize?: number;
  inputFontColor?: string;
  inputBoxColor?: string;
  presetRange?: { from: Date; to: Date } | null;
  defaultRange?: { from: Date; to: Date } | null;
}

const DateRangePickerWithSlider: React.FC<DateRangeSliderProps> = memo(
  ({
    minDate,
    maxDate,
    currentMinDate,
    currentMaxDate,
    onDateChange,
    formatDateForSlider,
    onOpenDialog,
    usePopupOnly = false,
    datePickerType = 'slider',
    inputFontSize,
    inputFontColor,
    inputBoxColor,
    presetRange,
    defaultRange,
  }) => {
    // Simple calculation: total days between dates
    const totalDays = Math.ceil(
      (maxDate.getTime() - minDate.getTime()) / (1000 * 60 * 60 * 24)
    );

    // Convert current dates to percentages (0-100) - much simpler
    const getPercentage = useCallback(
      (date: Date) => {
        const daysFromStart = Math.ceil(
          (date.getTime() - minDate.getTime()) / (1000 * 60 * 60 * 24)
        );
        const percentage = (daysFromStart / totalDays) * 100;
        // Allow up to 100% for max date, but cap at 100%
        return Math.max(0, Math.min(100, percentage));
      },
      [minDate, totalDays]
    );

    // Convert percentage back to date - much simpler
    const getDateFromPercentage = useCallback(
      (percentage: number) => {
        // Handle edge case where percentage is 100% or very close to it
        if (percentage >= 100) {
          return new Date(maxDate);
        }
        const daysFromStart = Math.round((percentage / 100) * totalDays);
        const resultDate = new Date(minDate);
        resultDate.setDate(resultDate.getDate() + daysFromStart);
        return resultDate;
      },
      [minDate, maxDate, totalDays]
    );

    // Current slider values as percentages - initialize with currentMinDate/currentMaxDate
    const [sliderValues, setSliderValues] = useState<[number, number]>([
      getPercentage(currentMinDate),
      getPercentage(currentMaxDate),
    ]);

    // Ref to track current values during drag
    const currentValuesRef = useRef<[number, number]>([
      getPercentage(currentMinDate),
      getPercentage(currentMaxDate),
    ]);

    // Dragging state
    const [isDragging, setIsDragging] = useState<number | null>(null);

    // Track if user has made changes to prevent override
    const [hasUserChanged, setHasUserChanged] = useState(false);
    const [isInitialized, setIsInitialized] = useState(false);

    // Initialize slider values only once
    useEffect(() => {
      if (!isInitialized) {
        const newMinPercent = getPercentage(currentMinDate);
        const newMaxPercent = getPercentage(currentMaxDate);
        const newValues: [number, number] = [newMinPercent, newMaxPercent];
        setSliderValues(newValues);
        currentValuesRef.current = newValues;
        setIsInitialized(true);
      }
    }, [currentMinDate, currentMaxDate, getPercentage, isInitialized]);

    // Handle external date updates (from calendar)
    useEffect(() => {
      if (isInitialized) {
        const newMinPercent = getPercentage(currentMinDate);
        const newMaxPercent = getPercentage(currentMaxDate);
        const newValues: [number, number] = [newMinPercent, newMaxPercent];
        setSliderValues(newValues);
        currentValuesRef.current = newValues;
      }
    }, [currentMinDate, currentMaxDate, getPercentage, isInitialized]);

    // Reset when min/max range changes (new preset)
    useEffect(() => {
      setHasUserChanged(false);
      setIsInitialized(false);
    }, [minDate, maxDate]);

    // Handle slider change - only update visual state, not filter
    const handleSliderChange = useCallback((values: [number, number]) => {
      const [minPercent, maxPercent] = values;

      // Ensure min is not greater than max
      const finalMinPercent = Math.min(minPercent, maxPercent);
      const finalMaxPercent = Math.max(minPercent, maxPercent);

      const newValues: [number, number] = [finalMinPercent, finalMaxPercent];
      setSliderValues(newValues);
      currentValuesRef.current = newValues;

      // Don't call onDateChange here - only update when drag stops
    }, []);

    // Create drag handler for thumbs
    const createDragHandler = useCallback(
      (thumbIndex: 0 | 1) => {
        return (e: React.MouseEvent) => {
          e.preventDefault();
          e.stopPropagation();

          setIsDragging(thumbIndex);

          const startX = e.clientX;
          const startValues = [...sliderValues] as [number, number];
          const trackElement = e.currentTarget.parentElement;

          if (!trackElement) return;

          const trackRect = trackElement.getBoundingClientRect();
          const trackWidth = trackRect.width;

          const handleMouseMove = (e: MouseEvent) => {
            const deltaX = e.clientX - startX;
            const deltaPercent = (deltaX / trackWidth) * 100;
            // Allow full range movement during dragging (0-100%)
            const newValue = Math.max(
              0,
              Math.min(100, startValues[thumbIndex] + deltaPercent)
            );

            const newValues = [...startValues] as [number, number];
            newValues[thumbIndex] = newValue;

            // Handle the case where both sliders are at the same position
            // Allow temporary overlap during dragging for better UX
            if (thumbIndex === 0) {
              // Left slider: ensure it doesn't go beyond right slider
              newValues[0] = Math.min(newValue, newValues[1]);
            } else {
              // Right slider: allow it to move left of left slider temporarily
              // This enables smooth dragging when both sliders start at the same position
              newValues[1] = newValue;
            }

            setSliderValues(newValues);
            currentValuesRef.current = newValues;
          };

          const handleMouseUp = () => {
            setIsDragging(null);
            setHasUserChanged(true); // Mark that user has made changes
            document.removeEventListener("mousemove", handleMouseMove);
            document.removeEventListener("mouseup", handleMouseUp);
            document.body.style.userSelect = "";
            document.body.style.cursor = "";

            // Use the ref to get the current values at the time of mouse up
            let currentValues = currentValuesRef.current;

            // Final positioning: ensure proper min/max relationship
            if (currentValues[0] > currentValues[1]) {
              // If left slider ended up to the right of right slider, swap them
              [currentValues[0], currentValues[1]] = [
                currentValues[1],
                currentValues[0],
              ];
              setSliderValues([...currentValues]);
              currentValuesRef.current = currentValues;
            }

            const newMinDate = getDateFromPercentage(currentValues[0]);
            const newMaxDate = getDateFromPercentage(currentValues[1]);
            onDateChange(newMinDate, newMaxDate);
          };

          // Prevent text selection during drag
          document.body.style.userSelect = "none";
          document.body.style.cursor = "grabbing";

          document.addEventListener("mousemove", handleMouseMove);
          document.addEventListener("mouseup", handleMouseUp);
        };
      },
      [getDateFromPercentage, onDateChange, sliderValues]
    );

    // Get dates for display
    const minDateDisplay = getDateFromPercentage(sliderValues[0]);
    const maxDateDisplay = getDateFromPercentage(sliderValues[1]);

    return (
      <div className="range-slider-container">
        {/* Add calendar picker as additional option */}
        <div
          style={{
            // marginTop: "5px",
            display: "flex",
            // justifyContent: "flex-start",
            // padding:"auto"
            justifyContent:"center"
          }}
        >
          <DateRangePicker
            from={currentMinDate}
            to={currentMaxDate}
            onOpenDialog={onOpenDialog}
            onDateChange={onDateChange}
            placeholder="Or select dates from calendar"
            minDate={minDate}
            maxDate={maxDate}
            inputFontSize={inputFontSize}
            inputFontColor={inputFontColor}
            inputBoxColor={inputBoxColor}
            preset={presetRange || undefined}
            defaultRange={defaultRange || undefined}
          />
        </div>
        {!usePopupOnly && datePickerType === 'slider' && (
          <div style={{ marginTop: "50px" }} className="simple-slider-wrapper">
            {/* Track */}
            <div className="simple-track">
              {/* Range fill */}
              <div
                className="simple-range-fill"
                style={{
                  left: `${sliderValues[0]}%`,
                  width: `${sliderValues[1] - sliderValues[0]}%`,
                }}
              />

              {/* Min thumb */}
              <div
                className="simple-thumb start-thumb"
                style={{ left: `${sliderValues[0]}%` }}
                onMouseDown={createDragHandler(0)}
              >
                <div className="thumb-label">
                  {formatDateForSlider(minDateDisplay)}
                </div>
              </div>

              {/* Max thumb */}
              <div
                className="simple-thumb end-thumb"
                style={{ left: `${sliderValues[1]}%` }}
                onMouseDown={createDragHandler(1)}
              >
                <div className="thumb-label">
                  {formatDateForSlider(maxDateDisplay)}
                </div>
              </div>
            </div>
          </div>
        )}
        {!usePopupOnly && datePickerType === 'calender' && (
          (() => {
            const [range, setRange] = useState<{ from?: Date; to?: Date } | undefined>({
              from: currentMinDate,
              to: currentMaxDate
            });
            const [fromText, setFromText] = useState<string>('');
            const [toText, setToText] = useState<string>('');
            const [month, setMonth] = useState<Date>(currentMaxDate);

            // Initialize text values
            React.useEffect(() => {
              setFromText(format(currentMinDate, "MM/dd/yyyy"));
              setToText(format(currentMaxDate, "MM/dd/yyyy"));
            }, [currentMinDate, currentMaxDate]);

            return React.createElement('div', {
              key: 'calendar',
              style: {
                display: 'flex',
                justifyContent: 'center',
                marginTop: '20px'
              }
            },
              React.createElement(Calendar, {
                mode: "range",
                selected: range?.from && range?.to ? range as DateRange : undefined,
                month: month,
                onMonthChange: setMonth,
                onSelect: (selectedRange: { from?: Date; to?: Date } | undefined) => {
                  console.log("Selected range:", selectedRange);
                  
                  if (selectedRange) {
                    const normalizedFrom = selectedRange.from || selectedRange.to;
                    const normalizedTo = selectedRange.to || selectedRange.from;
                    console.log("pickeddate from:", normalizedFrom);
                    console.log("pickeddate to:", normalizedTo);

                    // Update range state with the selected range
                    setRange({
                      from: normalizedFrom,
                      to: normalizedTo
                    });

                    if (normalizedFrom) {
                      setFromText(format(normalizedFrom, "MM/dd/yyyy"));
                    }
                    if (normalizedTo) {
                      setToText(format(normalizedTo, "MM/dd/yyyy"));
                    }

                    if (normalizedFrom && normalizedTo) {
                      const fromDate = minDate && normalizedFrom < minDate ? minDate : normalizedFrom;
                      const toDate = maxDate && normalizedTo > maxDate ? maxDate : normalizedTo;
                      onDateChange(fromDate, toDate);
                    }
                  } else {
                    setRange(undefined);
                    setFromText('');
                    setToText('');
                  }
                },
                numberOfMonths: 1
              })
            );
          })()
        )}


      </div>
    );
  }
);

export default DateRangePickerWithSlider;
