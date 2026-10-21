import { ChangeDetectionStrategy, Component, EventEmitter, OnInit, Output, signal } from '@angular/core';
import { CommonModule } from '@angular/common';
import { FormsModule } from '@angular/forms';

export interface FunctionDoc {
  name: string;
  category: string;
  syntax: string;
  description: string;
  sampleArgs: string;
}

@Component({
  selector: 'app-function-picker-dialog',
  standalone: true,
  imports: [CommonModule, FormsModule],
  templateUrl: './function-picker-dialog.component.html',
  styleUrls: ['./function-picker-dialog.component.css'],
  changeDetection: ChangeDetectionStrategy.OnPush,
})
export class FunctionPickerDialogComponent implements OnInit {
  @Output() readonly selectFunction = new EventEmitter<string>();
  @Output() readonly closeDialog = new EventEmitter<void>();

  readonly allFunctions: FunctionDoc[] = [
    {
      name: 'SUM',
      category: 'Math & Trig',
      syntax: 'SUM(number1, [number2], ...)',
      description: 'Adds all the numbers in a range of cells.',
      sampleArgs: 'B2:B4',
    },
    {
      name: 'AVERAGE',
      category: 'Statistical',
      syntax: 'AVERAGE(number1, [number2], ...)',
      description: 'Returns the average (arithmetic mean) of the arguments.',
      sampleArgs: 'B2:B4',
    },
    {
      name: 'COUNT',
      category: 'Statistical',
      syntax: 'COUNT(value1, [value2], ...)',
      description: 'Counts the number of cells that contain numbers.',
      sampleArgs: 'B2:B4',
    },
    {
      name: 'COUNTA',
      category: 'Statistical',
      syntax: 'COUNTA(value1, [value2], ...)',
      description: 'Counts the number of cells that are not empty.',
      sampleArgs: 'A2:A10',
    },
    {
      name: 'MAX',
      category: 'Statistical',
      syntax: 'MAX(number1, [number2], ...)',
      description: 'Returns the largest value in a set of values.',
      sampleArgs: 'B2:B4',
    },
    {
      name: 'MIN',
      category: 'Statistical',
      syntax: 'MIN(number1, [number2], ...)',
      description: 'Returns the smallest number in a set of values.',
      sampleArgs: 'B2:B4',
    },
    {
      name: 'ROUND',
      category: 'Math & Trig',
      syntax: 'ROUND(number, num_digits)',
      description: 'Rounds a number to a specified number of digits.',
      sampleArgs: 'B2, 2',
    },
    {
      name: 'IF',
      category: 'Logical',
      syntax: 'IF(logical_test, [value_if_true], [value_if_false])',
      description: 'Specifies a logical test to perform.',
      sampleArgs: 'B2>50, "Pass", "Fail"',
    },
    {
      name: 'SUMIF',
      category: 'Math & Trig',
      syntax: 'SUMIF(range, criteria, [sum_range])',
      description: 'Adds the cells specified by a given condition or criteria.',
      sampleArgs: 'A2:A10, "Pencil", B2:B10',
    },
    {
      name: 'COUNTIF',
      category: 'Statistical',
      syntax: 'COUNTIF(range, criteria)',
      description: 'Counts the number of cells within a range that meet the given condition.',
      sampleArgs: 'B2:B10, ">50"',
    },
    {
      name: 'AVERAGEIF',
      category: 'Statistical',
      syntax: 'AVERAGEIF(range, criteria, [average_range])',
      description: 'Returns the average (arithmetic mean) for cells specified by a condition.',
      sampleArgs: 'A2:A10, "Pencil", B2:B10',
    },
    {
      name: 'VLOOKUP',
      category: 'Lookup & Reference',
      syntax: 'VLOOKUP(lookup_value, table_array, col_index_num, [range_lookup])',
      description: 'Looks for a value in the leftmost column of a table, and then returns a value in the same row from a column you specify.',
      sampleArgs: 'A2, A2:C10, 2, FALSE',
    },
    {
      name: 'TEXT',
      category: 'Text',
      syntax: 'TEXT(value, format_text)',
      description: 'Converts a value to text in a specific number format.',
      sampleArgs: 'B2, "$#,##0.00"',
    },
    {
      name: 'CONCATENATE',
      category: 'Text',
      syntax: 'CONCATENATE(text1, [text2], ...)',
      description: 'Joins several text strings into one text string.',
      sampleArgs: 'A2, " ", B2',
    },
    {
      name: 'DATE',
      category: 'Date & Time',
      syntax: 'DATE(year, month, day)',
      description: 'Returns the serial number of a particular date.',
      sampleArgs: '2026, 9, 28',
    },
    {
      name: 'TODAY',
      category: 'Date & Time',
      syntax: 'TODAY()',
      description: 'Returns the current date formatted as a date.',
      sampleArgs: '',
    },
    {
      name: 'YEAR',
      category: 'Date & Time',
      syntax: 'YEAR(serial_number)',
      description: 'Returns the year of a date, an integer in the range 1900-9999.',
      sampleArgs: 'A2',
    },
    {
      name: 'MONTH',
      category: 'Date & Time',
      syntax: 'MONTH(serial_number)',
      description: 'Returns the month of a date, a number from 1 (January) to 12 (December).',
      sampleArgs: 'A2',
    },
    {
      name: 'DAY',
      category: 'Date & Time',
      syntax: 'DAY(serial_number)',
      description: 'Returns the day of the month, a number from 1 to 31.',
      sampleArgs: 'A2',
    },
  ];

  readonly selectedCategory = signal<string>('All');
  readonly selectedFunction = signal<FunctionDoc>(this.allFunctions[0]);
  readonly categories = ['All', 'Math & Trig', 'Statistical', 'Logical', 'Lookup & Reference', 'Text', 'Date & Time'];

  ngOnInit(): void {}

  get filteredFunctions(): FunctionDoc[] {
    const cat = this.selectedCategory();
    if (cat === 'All') return this.allFunctions;
    return this.allFunctions.filter((f) => f.category === cat);
  }

  onSelect(fn: FunctionDoc): void {
    this.selectedFunction.set(fn);
  }

  onApply(): void {
    const fn = this.selectedFunction();
    this.selectFunction.emit(fn.name);
    this.closeDialog.emit();
  }

  onCancel(): void {
    this.closeDialog.emit();
  }
}
