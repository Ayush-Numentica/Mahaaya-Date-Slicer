"use strict"

// Import necessary dependencies
import { format, parse, isValid } from "date-fns" // For date formatting and parsing
import { Calendar as CalendarIcon, Eraser as EraserIcon } from "lucide-react" // Icons
import { useState, useEffect } from "react" // For state management

// Import UI components

// Interface defining the props for DateRangePicker component
interface DateRangePickerProps {
  from?: Date // Start date of the range
  to?: Date // End date of the range
  onOpenDialog?: (fromDate: Date, toDate: Date, minDate: Date, maxDate: Date) => void // Callback function when dialog opens
  onDateChange?: (fromDate: Date | null, toDate: Date | null) => void // Callback function when dates change
  placeholder?: string // Placeholder text for inputs
  disabled?: boolean // Whether the component is disabled
  minDate?: Date // Minimum selectable date
  maxDate?: Date // Maximum selectable date
  // Optional preset and default range for clear behavior
  preset?: { from: Date; to: Date } | null
  defaultRange?: { from: Date; to: Date }
  // External control hooks
  onClear?: (range: { from: Date; to: Date }) => void
  resetSignal?: number | string // change this value externally to trigger clear
  registerClearHandler?: (handler: () => void) => void // parent can register a clear function
  inputFontSize?: number // font size for the input boxes (in px)
  inputFontColor?: string // font color for the input boxes (CSS color)
  inputBoxColor?: string
}

// Main DateRangePicker component function
export function DateRangePicker({
  from,
  to,
  onOpenDialog,
  onDateChange,
  placeholder = "Pick a date range", // Default placeholder text
  disabled = false, // Default disabled state
  minDate,
  maxDate,
  preset,
  defaultRange,
  onClear,
  resetSignal,
  registerClearHandler,
  inputFontSize,
  inputFontColor,
  inputBoxColor
}: DateRangePickerProps) {

  // Local state for input values
  const [fromInputValue, setFromInputValue] = useState<string>("");
  const [toInputValue, setToInputValue] = useState<string>("");
  const [showClear, setShowClear] = useState<boolean>(false);

  // Compute numeric font size with fallback
  const computedFontSize = typeof inputFontSize === 'number' && !isNaN(inputFontSize) && inputFontSize > 0
    ? `${inputFontSize}px`
    : "18px";
  const computedFontColor = inputFontColor && inputFontColor.trim() ? inputFontColor : '#000000';
  const computedBoxColor = inputBoxColor && inputBoxColor.trim() ? inputBoxColor : '#ffffff';

  // Clamp a date within optional [minDate, maxDate]
  const clampDate = (date: Date): Date => {
    let d = date;
    if (minDate && d < minDate) d = minDate;
    if (maxDate && d > maxDate) d = maxDate;
    return d;
  }

  // Resolve the range to use for clear
  const resolveClearRange = (): { from: Date; to: Date } | null => {
    let range: { from: Date; to: Date } | null = null;
    if (preset && preset.from && preset.to) {
      range = { from: preset.from, to: preset.to };
    } else if (defaultRange && defaultRange.from && defaultRange.to) {
      range = { from: defaultRange.from, to: defaultRange.to };
    } else if (minDate && maxDate) {
      range = { from: minDate, to: maxDate };
    }
    if (!range) return null;
    return { from: clampDate(range.from), to: clampDate(range.to) };
  }

  // Function to handle opening the Power BI date picker dialog
  const handleOpenDialog = () => {
    // Check if all required parameters are available before opening dialog
    if (onOpenDialog && from && to && minDate && maxDate) {
      onOpenDialog(from, to, minDate, maxDate);
    }
  }

  // Clear selection to preset/default range
  const handleClear = () => {
    const range = resolveClearRange();
    if (!range) return;
    if (onDateChange) {
      onDateChange(range.from, range.to);
    }
    if (onClear) {
      onClear(range);
    }
  }

  // Update input values when props change
  useEffect(() => {
    setFromInputValue(from ? format(from, "MM/dd/yyyy") : "");
  }, [from]);

  useEffect(() => {
    setToInputValue(to ? format(to, "MM/dd/yyyy") : "");
  }, [to]);

  // External reset trigger
  useEffect(() => {
    if (resetSignal === undefined) return;
    handleClear();
  }, [resetSignal]);

  // Register a programmatic clear handler
  useEffect(() => {
    if (!registerClearHandler) return;
    registerClearHandler(handleClear);
  }, [registerClearHandler, preset, defaultRange, minDate, maxDate]);

  // Function to parse date from input string
  const parseDate = (dateString: string): Date | null => {
    if (!dateString.trim()) return null;

    // Try parsing with dd/MM/yyyy format
    const parsedDate = parse(dateString, "MM/dd/yyyy", new Date());
    if (isValid(parsedDate)) {
      return parsedDate;
    }

    // Try parsing with other common formats
    const formats = ["dd-MM-yyyy", "dd.MM.yyyy", "yyyy-MM-dd", "MM/dd/yyyy"];
    for (const format of formats) {
      const parsed = parse(dateString, format, new Date());
      if (isValid(parsed)) {
        return parsed;
      }
    }

    return null;
  }

  // Check if string is a complete dd/MM/yyyy
  const isFullDDMMYYYY = (s: string): boolean => /^\d{2}\/\d{2}\/\d{4}$/.test(s);

  // Function to handle from date input change
  const handleFromDateChange = (event: React.ChangeEvent<HTMLInputElement>) => {
    const inputValue = event.target.value;
    setFromInputValue(inputValue); // Update local state immediately

    // Only propagate when a full valid date is present
    if (onDateChange && isFullDDMMYYYY(inputValue)) {
      const parsedDate = parseDate(inputValue);
      if (parsedDate) {
        const normalizedTo = to || parsedDate;
        onDateChange(parsedDate, normalizedTo);
      }
    }
  }

  // Function to handle to date input change
  const handleToDateChange = (event: React.ChangeEvent<HTMLInputElement>) => {
    const inputValue = event.target.value;
    setToInputValue(inputValue); // Update local state immediately

    // Only propagate when a full valid date is present
    if (onDateChange && isFullDDMMYYYY(inputValue)) {
      const parsedDate = parseDate(inputValue);
      if (parsedDate) {
        const normalizedFrom = from || parsedDate;
        onDateChange(normalizedFrom, parsedDate);
      }
    }
  }

  // Function to format the "from" date for display
  const formatFromDate = () => {
    if (from) {
      // If from date exists, format it as dd/MM/yyyy
      return format(from, "MM/dd/yyyy")
    } else {
      // If no from date, return empty string to show placeholder
      return ""
    }
  }

  // Function to format the "to" date for display
  const formatToDate = () => {
    if (to) {
      // If to date exists, format it as dd/MM/yyyy
      return format(to, "MM/dd/yyyy")
    } else {
      // If no to date, return empty string to show placeholder
      return ""
    }
  }

  // Return the JSX structure for the date range picker
  return (
    <div style={{ display: "flex", alignItems: 'center', width: '100%', justifyContent: 'space-between', paddingRight: "10px", gap: "4px", position: 'relative', maxWidth:"300px" }}
      onMouseEnter={() => setShowClear(true)}
      onMouseLeave={() => setShowClear(false)}
    >
      {/* Absolute Clear eraser at top-right of card */}
      <span
        role="button"
        aria-label="Clear selection"
        title="Clear selection"
        onClick={disabled ? undefined : handleClear}
        style={{ position: 'absolute', top: 2, right: 2, height: "10px", width: "10px", opacity: showClear ? (disabled ? 0.4 : 1) : 0, pointerEvents: showClear ? 'auto' : 'none', cursor: disabled ? 'not-allowed' : 'pointer', paddingRight: "10px", paddingBottom: "10px", display: 'none', alignItems: 'center', justifyContent: 'center' }}
      >
        <EraserIcon style={{ height: "15px", width: "15px", marginBottom:"10px"}} />
      </span>
      <div style={{ display: 'flex', alignItems: 'center', gap: '2px', width: '100%', minWidth: '100px', maxWidth: '300px', paddingLeft: "10px", paddingTop: "10px", paddingBottom: "10px" }}>
        {/* First input box for "From Date" */}
        <input
          type="text"
          value={fromInputValue} // Use local state for input value
          onChange={handleFromDateChange} // Handle date input changes
          disabled={disabled} // Respect disabled prop
          placeholder="From Date" // Placeholder text when empty
          title={fromInputValue || "From Date"} // Tooltip showing full content on hover
          className={`input-box`}
          style={{
            flex: '1',
            minWidth: "50px",
            maxWidth: '120px',
            outline: 'none',
            textOverflow: 'ellipsis',
            overflow: 'hidden',
            whiteSpace: 'nowrap',
            height: computedFontSize,
            // minHeight: '32px',
            resize: 'none',
            border: "1px solid #ccc",
            padding: "4px",
            borderRadius: "4px",
            fontSize: computedFontSize
            , color: computedFontColor,
            backgroundColor: computedBoxColor
          }} // Input box with ellipsis overflow
        />

        {/* Visual separator between the two date inputs */}
        <span style={{ color: '#666', fontSize: '14px', flexShrink: 0 }}>-</span>

        {/* Second input box for "To Date" */}
        <input
          type="text"
          value={toInputValue} // Use local state for input value
          onChange={handleToDateChange} // Handle date input changes
          disabled={disabled} // Respect disabled prop
          placeholder="To Date" // Placeholder text when empty
          title={toInputValue || "To Date"} // Tooltip showing full content on hover
          className={`input-box`}
          style={{
            flex: '1',
            minWidth: "50px",
            maxWidth: '120px',
            outline: 'none',
            textOverflow: 'ellipsis',
            overflow: 'hidden',
            whiteSpace: 'nowrap',
            height: computedFontSize,
            // minHeight: '32px',
            resize: 'none',
            border: "1px solid #ccc",
            padding: "4px",
            borderRadius: "4px",
            fontSize: computedFontSize
            , color: computedFontColor,
            backgroundColor: computedBoxColor
          }} // Input box with ellipsis overflow
        />
      </div>
      <div style={{ display: 'flex', flexDirection: 'column', alignItems: 'center', gap: '6px', paddingRight: "4px" }}>
        {/* Calendar button to open date picker dialog */}
        <CalendarIcon
          style={{ height: "18px", width: "18px", cursor: disabled ? 'not-allowed' : 'pointer', opacity: disabled ? 0.5 : 1 }}
          onClick={disabled ? undefined : handleOpenDialog}
        /> {/* Calendar icon */}
      </div>

    </div>
  )
}
