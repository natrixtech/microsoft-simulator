import { ActivityPayload } from 'shared';

export const EXCEL_LESSONS: Record<string, ActivityPayload> = {
  cellValue: {
    app: 'excel',
    mode: 'lesson',
    title: 'Enter Data into Cells',
    tasks: [
      {
        title: 'Type values into cells',
        explanation: 'Cells can store numbers and text. Click on a cell and type to enter data.',
        instructions: ['Click cell A2 and type `Pencil`.', 'Click cell B2 and type `1.50`.', 'Click Check My Work.'],
        hint: 'Type Pencil in cell A2 and 1.50 in cell B2.',
        startingSheet: {
          cells: {
            A1: 'Item',
            B1: 'Price',
          },
          formats: {
            'A1:B1': { bold: true },
          },
        },
        goals: [
          {
            type: 'cellValue',
            cells: {
              A2: 'Pencil',
              B2: 1.5,
            },
          },
        ],
        showMeHow: true,
      },
    ],
    assetBaseUrl: './assets/',
    theme: 'LIGHT',
  },

  cellFormula: {
    app: 'excel',
    mode: 'lesson',
    title: 'Calculate Sum with Formula',
    tasks: [
      {
        title: 'Enter SUM formula',
        explanation: 'Formulas start with an equals sign (=) and calculate values automatically.',
        instructions: ['Select cell B5.', 'Enter `=SUM(B2:B4)` into cell B5.', 'Press Enter.'],
        hint: 'Click B5 and type =SUM(B2:B4) then press Enter.',
        startingSheet: {
          cells: {
            A1: 'Item',
            B1: 'Sales',
            A2: 'Apples',
            B2: 15,
            A3: 'Oranges',
            B3: 25,
            A4: 'Bananas',
            B4: 10,
          },
          formats: {
            'A1:B1': { bold: true },
          },
        },
        goals: [
          {
            type: 'cellFormula',
            cell: 'B5',
            function: 'SUM',
            references: 'B2:B4',
          },
        ],
        showMeHow: true,
      },
    ],
    assetBaseUrl: './assets/',
    theme: 'LIGHT',
  },

  cellFormat: {
    app: 'excel',
    mode: 'lesson',
    title: 'Format a Range as Currency',
    tasks: [
      {
        title: 'Format a range as currency',
        explanation: 'Currency formatting tells the spreadsheet to display numbers with a dollar sign and two decimal places so prices are easy to scan.',
        instructions: ['Select B2:B4.', 'Click the $ button in the toolbar.'],
        hint: 'Drag from B2 down to B4, then click $.',
        startingSheet: {
          cells: {
            A1: 'Item',
            B1: 'Price',
            A2: 'Pencil',
            B2: 0.5,
            A3: 'Notebook',
            B3: 2.25,
            A4: 'Backpack',
            B4: 24,
            B5: '=SUM(B2:B4)',
          },
          formats: {
            'A1:B1': { bold: true },
          },
        },
        goals: [
          {
            type: 'cellFormat',
            range: 'B2:B4',
            numberFormat: 'currency',
          },
        ],
        showMeHow: true,
      },
    ],
    assetBaseUrl: './assets/',
    theme: 'LIGHT',
  },

  sorted: {
    app: 'excel',
    mode: 'lesson',
    title: 'Sort Data by Price',
    tasks: [
      {
        title: 'Sort range descending',
        explanation: 'Sorting helps organize data so you can quickly find highest or lowest values.',
        instructions: ['Select range A1:B5.', 'Click Sort in the Data tab and sort by Column B descending (Z to A).'],
        hint: 'Go to Data tab, open Sort, select Column B, choose Z to A, and click OK.',
        startingSheet: {
          cells: {
            A1: 'Item',
            B1: 'Price',
            A2: 'Pen',
            B2: 2,
            A3: 'Desk',
            B3: 150,
            A4: 'Chair',
            B4: 75,
            A5: 'Lamp',
            B5: 35,
          },
          formats: {
            'A1:B1': { bold: true },
          },
        },
        goals: [
          {
            type: 'sorted',
            range: 'A1:B5',
            byColumn: 'B',
            order: 'desc',
            hasHeader: true,
          },
        ],
        showMeHow: true,
      },
    ],
    assetBaseUrl: './assets/',
    theme: 'LIGHT',
  },

  filter: {
    app: 'excel',
    mode: 'lesson',
    title: 'Filter High Value Orders',
    tasks: [
      {
        title: 'Create a filter',
        explanation: 'Filters allow you to temporarily hide rows that do not match specified criteria.',
        instructions: ['Select range A1:C5.', 'Click Filter on the Data tab.', 'Filter Column C for values greater than 50.'],
        hint: 'Select A1:C5, click Filter in Data tab, then open Column C filter and set condition to greaterThan 50.',
        startingSheet: {
          cells: {
            A1: 'Order',
            B1: 'Customer',
            C1: 'Total',
            A2: '101',
            B2: 'Acme',
            C2: 120,
            A3: '102',
            B3: 'Beta',
            C3: 40,
            A4: '103',
            B4: 'Gamma',
            C4: 85,
            A5: '104',
            B5: 'Delta',
            C5: 30,
          },
          formats: {
            'A1:C1': { bold: true },
          },
        },
        goals: [
          {
            type: 'filter',
            range: 'A1:C5',
            column: 'C',
            condition: 'greaterThan',
            value: 50,
          },
        ],
        showMeHow: true,
      },
    ],
    assetBaseUrl: './assets/',
    theme: 'LIGHT',
  },

  chart: {
    app: 'excel',
    mode: 'lesson',
    title: 'Insert Column Chart',
    tasks: [
      {
        title: 'Insert a chart',
        explanation: 'Charts provide visual summaries of spreadsheet numbers.',
        instructions: ['Select range A1:B5.', 'Go to Insert tab, click Column Chart, and set title to `Sales Report`.'],
        hint: 'Select A1:B5, click Column Chart on Insert tab, enter title Sales Report and insert.',
        startingSheet: {
          cells: {
            A1: 'Month',
            B1: 'Revenue',
            A2: 'Jan',
            B2: 12000,
            A3: 'Feb',
            B3: 15000,
            A4: 'Mar',
            B4: 18000,
            A5: 'Apr',
            B5: 14000,
          },
          formats: {
            'A1:B1': { bold: true },
          },
        },
        goals: [
          {
            type: 'chart',
            chartType: 'column',
            range: 'A1:B5',
            title: 'Sales Report',
          },
        ],
        showMeHow: true,
      },
    ],
    assetBaseUrl: './assets/',
    theme: 'LIGHT',
  },

  pivotTable: {
    app: 'excel',
    mode: 'lesson',
    title: 'Insert PivotTable',
    tasks: [
      {
        title: 'Insert a pivot table',
        explanation: 'PivotTables summarize large datasets by grouping categories and calculating totals.',
        instructions: ['Select range A1:C6.', 'Click PivotTable on the Insert tab.', 'Set Rows to `Category` and Values to `Amount` with SUM.'],
        hint: 'Open Insert tab > PivotTable, pick Category for Rows and Amount for Values, then click OK.',
        startingSheet: {
          cells: {
            A1: 'Item',
            B1: 'Category',
            C1: 'Amount',
            A2: 'Desk',
            B2: 'Office',
            C2: 200,
            A3: 'Pen',
            B3: 'Supplies',
            C3: 10,
            A4: 'Chair',
            B4: 'Office',
            C4: 150,
            A5: 'Paper',
            B5: 'Supplies',
            C5: 25,
            A6: 'Laptop',
            B6: 'Hardware',
            C6: 900,
          },
          formats: {
            'A1:C1': { bold: true },
          },
        },
        goals: [
          {
            type: 'pivotTable',
            range: 'A1:C6',
            rowsField: 'Category',
            valuesField: 'Amount',
            summarize: 'SUM',
          },
        ],
        showMeHow: true,
      },
    ],
    assetBaseUrl: './assets/',
    theme: 'LIGHT',
  },

  comment: {
    app: 'excel',
    mode: 'lesson',
    title: 'Add a Cell Comment',
    tasks: [
      {
        title: 'Add a comment',
        explanation: 'Comments let you leave notes and feedback attached to specific cells.',
        instructions: ['Select cell C4.', 'Click Comment on the Insert tab and type `check pricing`.'],
        hint: 'Select C4, click Comment on Insert tab, type check pricing and post.',
        startingSheet: {
          cells: {
            A1: 'Product',
            B1: 'Cost',
            C1: 'Retail',
            A2: 'Widget',
            B2: 5,
            C2: 12,
            A3: 'Gizmo',
            B3: 8,
            C3: 19,
            A4: 'Doohickey',
            B4: 15,
            C4: 22,
          },
          formats: {
            'A1:C1': { bold: true },
          },
        },
        goals: [
          {
            type: 'comment',
            cell: 'C4',
            contains: 'check pricing',
          },
        ],
        showMeHow: true,
      },
    ],
    assetBaseUrl: './assets/',
    theme: 'LIGHT',
  },

  conditionalFormat: {
    app: 'excel',
    mode: 'lesson',
    title: 'Highlight Values with Conditional Formatting',
    tasks: [
      {
        title: 'Add conditional format rule',
        explanation: 'Conditional formatting changes the appearance of cells based on rules you set.',
        instructions: ['Select range C2:C8.', 'Click Conditional Formatting on Home tab.', 'Highlight cells greater than 100 with light green.'],
        hint: 'Select C2:C8, open Conditional Formatting on Home tab, pick greaterThan 100 with lightGreen fill.',
        startingSheet: {
          cells: {
            A1: 'ID',
            B1: 'Name',
            C1: 'Score',
            A2: '1',
            B2: 'Alex',
            C2: 125,
            A3: '2',
            B3: 'Blake',
            C3: 90,
            A4: '3',
            B4: 'Casey',
            C4: 110,
            A5: '4',
            B5: 'Dana',
            C5: 85,
            A6: '5',
            B6: 'Evan',
            C6: 150,
            A7: '6',
            B7: 'Fox',
            C7: 75,
            A8: '7',
            B8: 'Glen',
            C8: 105,
          },
          formats: {
            'A1:C1': { bold: true },
          },
        },
        goals: [
          {
            type: 'conditionalFormat',
            range: 'C2:C8',
            condition: 'greaterThan',
            value: 100,
            fillColor: 'lightGreen',
          },
        ],
        showMeHow: true,
      },
    ],
    assetBaseUrl: './assets/',
    theme: 'LIGHT',
  },

  dataValidation: {
    app: 'excel',
    mode: 'lesson',
    title: 'Create Dropdown with Data Validation',
    tasks: [
      {
        title: 'Add list data validation',
        explanation: 'Data validation restricts the type of data or values that users can enter into a cell.',
        instructions: ['Select range B2:B6.', 'Click Data Validation on Data tab.', 'Set allowed items to `Yes, No`.'],
        hint: 'Select B2:B6, open Data Validation on Data tab, enter Yes, No and click OK.',
        startingSheet: {
          cells: {
            A1: 'Employee',
            B1: 'Completed Training',
            A2: 'Jordan',
            B2: 'Yes',
            A3: 'Taylor',
            B3: 'No',
            A4: 'Morgan',
            B4: 'Yes',
            A5: 'Sam',
            B5: '',
            A6: 'Riley',
            B6: '',
          },
          formats: {
            'A1:B1': { bold: true },
          },
        },
        goals: [
          {
            type: 'dataValidation',
            range: 'B2:B6',
            allowed: ['Yes', 'No'],
          },
        ],
        showMeHow: true,
      },
    ],
    assetBaseUrl: './assets/',
    theme: 'LIGHT',
  },

  noDuplicates: {
    app: 'excel',
    mode: 'lesson',
    title: 'Remove Duplicate Rows',
    tasks: [
      {
        title: 'Remove duplicates',
        explanation: 'Remove duplicates scans a range and discards duplicate rows while retaining unique records.',
        instructions: ['Select range A1:C6.', 'Click Remove Duplicates on Data tab.', 'Ensure Data has headers is checked and remove duplicates.'],
        hint: 'Select A1:C6, click Remove Duplicates in Data tab and confirm.',
        startingSheet: {
          cells: {
            A1: 'First',
            B1: 'Last',
            C1: 'City',
            A2: 'John',
            B2: 'Doe',
            C2: 'Dallas',
            A3: 'Jane',
            B3: 'Smith',
            C3: 'Austin',
            A4: 'John',
            B4: 'Doe',
            C4: 'Dallas',
            A5: 'Mark',
            B5: 'Taylor',
            C5: 'Houston',
            A6: 'Jane',
            B6: 'Smith',
            C6: 'Austin',
          },
          formats: {
            'A1:C1': { bold: true },
          },
        },
        goals: [
          {
            type: 'noDuplicates',
            range: 'A2:C6',
          },
        ],
        showMeHow: true,
      },
    ],
    assetBaseUrl: './assets/',
    theme: 'LIGHT',
  },

  action: {
    app: 'excel',
    mode: 'lesson',
    title: 'Select a Range of Cells',
    tasks: [
      {
        title: 'Select cells B2 to B4',
        explanation: 'Selecting a range allows you to format or copy multiple cells at once.',
        instructions: ['Click on cell B2 and drag down to B4 to select B2:B4.', 'Click Check My Work.'],
        hint: 'Drag your mouse from cell B2 down to cell B4.',
        startingSheet: {
          cells: {
            A1: 'Item',
            B1: 'Qty',
            A2: 'Pencil',
            B2: 10,
            A3: 'Ruler',
            B3: 5,
            A4: 'Eraser',
            B4: 12,
          },
          formats: {
            'A1:B1': { bold: true },
          },
        },
        goals: [
          {
            type: 'action',
            action: 'selectRange',
            range: 'B2:B4',
          },
        ],
        showMeHow: true,
      },
    ],
    assetBaseUrl: './assets/',
    theme: 'LIGHT',
  },

  assessment: {
    app: 'excel',
    mode: 'assessment',
    title: 'Excel Skills Assessment',
    tasks: [
      {
        title: 'Enter Price',
        explanation: 'Enter the missing price for the notebook.',
        instructions: ['Click cell B3 and type `3.50`.', 'Click Check My Work.'],
        hint: 'Type 3.5 in cell B3.',
        startingSheet: {
          cells: {
            A1: 'Item',
            B1: 'Price',
            A2: 'Pencil',
            B2: 1.0,
            A3: 'Notebook',
            B3: '',
            A4: 'Binder',
            B4: 5.5,
          },
          formats: {
            'A1:B1': { bold: true },
          },
        },
        goals: [
          {
            type: 'cellValue',
            cells: {
              B3: 3.5,
            },
          },
        ],
      },
      {
        title: 'Calculate Total',
        explanation: 'Use the SUM function to calculate the total price.',
        instructions: ['Select cell B5.', 'Enter `=SUM(B2:B4)`.', 'Click Check My Work.'],
        hint: 'Enter =SUM(B2:B4) in cell B5.',
        startingSheet: null, // Inherits state from Task 1!
        goals: [
          {
            type: 'cellFormula',
            cell: 'B5',
            function: 'SUM',
            references: 'B2:B4',
          },
        ],
      },
      {
        title: 'Format Currency',
        explanation: 'Format all prices and the total as currency.',
        instructions: ['Select range B2:B5.', 'Click the $ button on the toolbar.', 'Click Check My Work.'],
        hint: 'Highlight B2:B5 and click $.',
        startingSheet: null, // Inherits state from Tasks 1 & 2!
        goals: [
          {
            type: 'cellFormat',
            range: 'B2:B5',
            numberFormat: 'currency',
          },
        ],
      },
    ],
    assetBaseUrl: './assets/',
    theme: 'LIGHT',
  },
};
