import React from 'react';
import { createRoot } from 'react-dom/client';
import DateRangePickerWithSlider from './DateRangePickerWithSlider';

interface ReactSliderWrapperProps {
  minDate: Date;
  maxDate: Date;
  currentMinDate: Date;
  currentMaxDate: Date;
  onDateChange: (minDate: Date, maxDate: Date) => void;
  formatDateForSlider: (date: Date) => string;
  onOpenDialog: (fromDate: Date, toDate: Date, minDate: Date, maxDate: Date) => void;
  usePopupOnly?: boolean;
  container: HTMLElement;
  datePickerType?: any;
  inputFontSize?: number;
  inputFontColor?: string;
  inputBoxColor?: string;
  presetRange?: { from: Date; to: Date } | null;
  defaultRange?: { from: Date; to: Date } | null;
}

export class ReactSliderWrapper {
  private root: any;
  private container: HTMLElement;
  private currentProps: Omit<ReactSliderWrapperProps, 'container'> | null = null;

  constructor(container: HTMLElement) {
    this.container = container;
    this.root = createRoot(container);
  }

  render(props: Omit<ReactSliderWrapperProps, 'container'>) {
    this.currentProps = props;
    console.log('props', props);
    this.root.render(
      React.createElement(DateRangePickerWithSlider, props)
    );
  }

  updateDates(minDate: Date, maxDate: Date, presetRange?: { from: Date; to: Date } | null) {
    if (this.currentProps) {
      const updatedProps = {
        ...this.currentProps,
        currentMinDate: minDate,
        currentMaxDate: maxDate,
        ...(presetRange !== undefined && { presetRange })
      };
      this.currentProps = updatedProps;
      console.log('updated props', updatedProps);
      this.root.render(
        React.createElement(DateRangePickerWithSlider, updatedProps)
      );
    }
  }

  destroy() {
    if (this.root) {
      this.root.unmount();
    }
  }
}
