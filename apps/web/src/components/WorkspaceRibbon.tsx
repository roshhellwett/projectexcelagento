import React, { useMemo, useRef, useState } from 'react';
import {
  AlignCenter,
  AlignLeft,
  AlignRight,
  BarChart3,
  Bold,
  Calculator,
  Check,
  CheckSquare,
  Clipboard,
  Code2,
  Columns3,
  Copy,
  Database,
  Download,
  Eraser,
  FileClock,
  FileSpreadsheet,
  Filter,
  FolderOpen,
  FunctionSquare,
  Grid2X2,
  HelpCircle,
  Highlighter,
  Image,
  Italic,
  LockKeyhole,
  MessageSquare,
  MoreHorizontal,
  MoveHorizontal,
  Paintbrush,
  PanelTop,
  PenLine,
  PieChart,
  Plus,
  Presentation,
  RefreshCw,
  RotateCcw,
  Scissors,
  SearchCheck,
  ShieldCheck,
  Sigma,
  SlidersHorizontal,
  SortAsc,
  Table2,
  Tags,
  Target,
  TextCursorInput,
  TextSelect,
  Underline,
  Upload,
  WandSparkles,
  WrapText,
  X,
} from 'lucide-react';

export type RibbonTabId =
  'home' | 'insert' | 'draw' | 'page-layout' | 'formulas' | 'data' | 'review' | 'automate' | 'help';

interface RibbonAction {
  id: string;
  label: string;
  icon: React.ComponentType<{ size?: number; strokeWidth?: number }>;
  hint?: string;
  tone?: 'default' | 'primary' | 'warning';
  disabled?: boolean;
}

interface RibbonGroup {
  label: string;
  actions: RibbonAction[];
}

interface WorkspaceRibbonProps {
  activeSheetName: string;
  isProcessing: boolean;
  onManualAction: (actionId: string) => void;
  onOpenOperationModal: () => void;
  onUploadFile: (file: File) => void;
  onExport: () => void;
  onUndo: () => void;
  onRedo: () => void;
  canUndo: boolean;
  canRedo: boolean;
}

const tabs: Array<{ id: RibbonTabId; label: string }> = [
  { id: 'home', label: 'Home' },
  { id: 'insert', label: 'Insert' },
  { id: 'draw', label: 'Draw' },
  { id: 'page-layout', label: 'Page Layout' },
  { id: 'formulas', label: 'Formulas' },
  { id: 'data', label: 'Data' },
  { id: 'review', label: 'Review' },
  { id: 'automate', label: 'Automate' },
  { id: 'help', label: 'Help' },
];

const command = (
  id: string,
  label: string,
  icon: RibbonAction['icon'],
  description: string,
  hint?: string,
  tone?: RibbonAction['tone'],
): RibbonAction => ({ id, label, icon, hint: hint ?? description, tone });

const comingSoon = (
  id: string,
  label: string,
  icon: RibbonAction['icon'],
  hint: string,
): RibbonAction => ({ id, label, icon, hint, disabled: true });

function actionGroups(tab: RibbonTabId, sheet: string): RibbonGroup[] {
  const useSheet = (instruction: string) => `${instruction} on the active sheet "${sheet}".`;
  switch (tab) {
    case 'home':
      return [
        {
          label: 'Clipboard',
          actions: [
            command(
              'paste',
              'Paste',
              Clipboard,
              'Paste the current clipboard contents into the selected cell.',
            ),
            command(
              'cut',
              'Cut',
              Scissors,
              'Cut the selected cells so they can be pasted elsewhere.',
            ),
            command('copy', 'Copy', Copy, 'Copy the selected cells for reuse.'),
            command(
              'format-painter',
              'Format Painter',
              Paintbrush,
              useSheet('Copy the formatting pattern from the selected area to a matching area'),
            ),
          ],
        },
        {
          label: 'Font',
          actions: [
            comingSoon(
              'font',
              'Font',
              TextCursorInput,
              'Font-family editing needs workbook style metadata.',
            ),
            comingSoon(
              'font-size',
              'Size',
              TextSelect,
              'Font-size editing needs workbook style metadata.',
            ),
            command('bold', 'Bold', Bold, useSheet('Make the selected cells bold')),
            command('italic', 'Italic', Italic, useSheet('Make the selected cells italic')),
            command('underline', 'Underline', Underline, useSheet('Underline the selected cells')),
            command(
              'fill',
              'Fill',
              Highlighter,
              useSheet('Apply a clear highlight fill to the selected cells'),
            ),
            command(
              'borders',
              'Borders',
              Grid2X2,
              useSheet('Apply appropriate borders to the selected cells'),
            ),
            command(
              'font-color',
              'Font Color',
              Paintbrush,
              useSheet('Apply a readable font color to the selected cells'),
            ),
          ],
        },
        {
          label: 'Alignment',
          actions: [
            command(
              'align-left',
              'Left',
              AlignLeft,
              useSheet('Align the selected cells to the left'),
            ),
            command(
              'align-center',
              'Center',
              AlignCenter,
              useSheet('Center-align the selected cells'),
            ),
            command(
              'align-right',
              'Right',
              AlignRight,
              useSheet('Align the selected cells to the right'),
            ),
            command(
              'orientation',
              'Orientation',
              MoveHorizontal,
              useSheet('Choose the best text orientation for the selected cells'),
            ),
            command('wrap', 'Wrap Text', WrapText, useSheet('Wrap text in the selected cells')),
            command(
              'merge',
              'Merge & Center',
              MoveHorizontal,
              useSheet('Merge the selected range and center its content'),
            ),
            command(
              'indent',
              'Indent',
              MoveHorizontal,
              useSheet('Indent the selected cells for clearer hierarchy'),
            ),
          ],
        },
        {
          label: 'Number',
          actions: [
            command(
              'general',
              'General',
              SlidersHorizontal,
              useSheet('Use General number formatting for the selected cells'),
            ),
            command(
              'currency',
              'Currency',
              Calculator,
              useSheet('Format the selected numeric cells as currency'),
            ),
            command(
              'percent',
              'Percent',
              Tags,
              useSheet('Format the selected numeric cells as percentages'),
            ),
            command(
              'comma',
              'Comma',
              MoreHorizontal,
              useSheet('Format the selected numbers with thousands separators'),
            ),
            command(
              'decimals',
              'Decimals',
              Plus,
              useSheet('Increase the displayed decimal places for the selected cells'),
            ),
          ],
        },
        {
          label: 'Styles',
          actions: [
            command(
              'conditional-formatting',
              'Conditional Formatting',
              FlagIcon,
              useSheet(
                'Suggest and apply useful conditional formatting rules to the selected range',
              ),
            ),
            command(
              'format-table',
              'Format as Table',
              Table2,
              useSheet('Convert the current data range into a clearly structured table'),
            ),
            command(
              'cell-styles',
              'Cell Styles',
              PanelTop,
              useSheet('Suggest a consistent professional cell style for this range'),
            ),
          ],
        },
        {
          label: 'Cells & Editing',
          actions: [
            command(
              'insert',
              'Insert',
              Plus,
              useSheet('Insert the requested rows or columns near the selection'),
            ),
            command(
              'delete',
              'Delete',
              X,
              useSheet('Delete the selected rows or columns after reviewing the impact'),
              undefined,
              'warning',
            ),
            command(
              'format',
              'Format',
              SlidersHorizontal,
              useSheet('Inspect and format the selected cells'),
            ),
            command(
              'autosum',
              'AutoSum',
              Sigma,
              useSheet('Add a verified sum for the most relevant numeric column'),
            ),
            command(
              'fill-series',
              'Fill',
              WandSparkles,
              useSheet('Fill the selected range using the best safe series or pattern'),
            ),
            command(
              'clear',
              'Clear',
              Eraser,
              useSheet('Clear the selected cells after showing what will be removed'),
              undefined,
              'warning',
            ),
            command(
              'sort-filter',
              'Sort & Filter',
              Filter,
              useSheet('Help me sort or filter this data based on the selection'),
            ),
            command(
              'find-select',
              'Find & Select',
              SearchCheck,
              'Open search and selection tools for this workbook.',
            ),
            command(
              'sensitivity',
              'Sensitivity',
              ShieldCheck,
              'Explain workbook sensitivity and available protection controls.',
            ),
            command('add-ins', 'Add-ins', Plus, 'Show the available workspace and agent add-ins.'),
          ],
        },
      ];
    case 'insert':
      return [
        {
          label: 'Tables',
          actions: [
            command(
              'pivot',
              'PivotTable',
              Grid2X2,
              useSheet('Create a pivot-style summary grouped by the relevant fields'),
            ),
            command(
              'recommended-pivot',
              'Recommended PivotTables',
              WandSparkles,
              useSheet('Recommend and create the most useful pivot-style summaries'),
            ),
            command(
              'table',
              'Table',
              Table2,
              useSheet('Create a structured table from the current data'),
            ),
            command(
              'forms',
              'Forms',
              CheckSquare,
              useSheet('Create a simple input form for this table'),
            ),
          ],
        },
        {
          label: 'Illustrations',
          actions: [
            comingSoon(
              'pictures',
              'Pictures',
              Image,
              'Image insertion needs a workbook drawing layer.',
            ),
            comingSoon(
              'shapes',
              'Shapes',
              Target,
              'Shape insertion needs a workbook drawing layer.',
            ),
            comingSoon(
              'icons',
              'Icons',
              Presentation,
              'Icon insertion needs a workbook drawing layer.',
            ),
            comingSoon(
              '3d-models',
              '3D Models',
              Presentation,
              '3D model insertion needs a workbook drawing layer.',
            ),
            comingSoon(
              'smartart',
              'SmartArt',
              Presentation,
              'SmartArt insertion needs a workbook drawing layer.',
            ),
            command(
              'screenshot',
              'Screenshot',
              Image,
              'Capture a screenshot of the current workspace for the agent to inspect.',
            ),
          ],
        },
        {
          label: 'Controls',
          actions: [
            command(
              'checkbox',
              'Checkbox',
              CheckSquare,
              useSheet('Add checkbox-style TRUE/FALSE controls to the selected cells'),
            ),
          ],
        },
        {
          label: 'Charts',
          actions: [
            command(
              'recommended-chart',
              'Recommended Charts',
              WandSparkles,
              useSheet('Choose the clearest chart for this data and prepare it'),
            ),
            command(
              'column-chart',
              'Column',
              BarChart3,
              useSheet('Create a column chart from the selected data'),
            ),
            command(
              'line-chart',
              'Line',
              Presentation,
              useSheet('Create a line chart from the selected data'),
            ),
            command(
              'pie-chart',
              'Pie',
              PieChart,
              useSheet('Create a pie chart from the selected data'),
            ),
            command(
              'scatter-chart',
              'Scatter',
              Target,
              useSheet('Create a scatter chart from the selected numeric columns'),
            ),
            command(
              'pivot-chart',
              'PivotChart',
              BarChart3,
              useSheet('Create a chart from a pivot-style summary'),
            ),
            command(
              'maps',
              'Maps',
              Target,
              useSheet('Create a map-ready summary from geography fields'),
            ),
          ],
        },
        {
          label: 'Sparklines & Filters',
          actions: [
            comingSoon(
              'sparklines',
              'Sparklines',
              BarChart3,
              'Sparkline drawing needs a chart artifact layer.',
            ),
            comingSoon(
              'sparkline-line',
              'Line',
              Presentation,
              'Line sparklines need a chart artifact layer.',
            ),
            comingSoon(
              'sparkline-column',
              'Column',
              BarChart3,
              'Column sparklines need a chart artifact layer.',
            ),
            comingSoon(
              'sparkline-win-loss',
              'Win/Loss',
              BarChart3,
              'Win/Loss sparklines need a chart artifact layer.',
            ),
            command(
              'slicer',
              'Slicer',
              Filter,
              useSheet('Create a filter view for the selected table'),
            ),
            command(
              'timeline',
              'Timeline',
              FileClock,
              useSheet('Create a time-based filter for the selected date column'),
            ),
          ],
        },
        {
          label: 'Links, Comments & Text',
          actions: [
            comingSoon(
              'link',
              'Link',
              MessageSquare,
              'Hyperlink preservation needs workbook metadata support.',
            ),
            command(
              'comment',
              'Comment',
              MessageSquare,
              'Add a note to the selected cell for the agent to remember.',
            ),
            comingSoon(
              'text-box',
              'Text Box',
              TextCursorInput,
              'Text boxes need a workbook drawing layer.',
            ),
            comingSoon(
              'header-footer',
              'Header & Footer',
              PanelTop,
              'Page headers and footers need print-layout metadata.',
            ),
            comingSoon(
              'word-art',
              'WordArt',
              TextCursorInput,
              'WordArt needs a workbook drawing layer.',
            ),
            comingSoon(
              'signature',
              'Signature Line',
              PenLine,
              'Signature lines need workbook drawing metadata.',
            ),
            comingSoon(
              'object',
              'Object',
              FileSpreadsheet,
              'Embedded objects need workbook package support.',
            ),
            comingSoon(
              'equation',
              'Equation',
              FunctionSquare,
              'Equation objects need a workbook drawing layer.',
            ),
            comingSoon(
              'symbol',
              'Symbol',
              TextCursorInput,
              'Symbol insertion is available through direct cell editing.',
            ),
          ],
        },
      ];
    case 'draw':
      return [
        {
          label: 'History',
          actions: [
            command('draw-undo', 'Undo', RotateCcw, 'Undo the last verified workbook change.'),
            command('draw-redo', 'Redo', RefreshCw, 'Redo the last undone workbook change.'),
          ],
        },
        {
          label: 'Drawing Tools',
          actions: [
            comingSoon('select', 'Select', Target, 'Drawing selection needs a workbook ink layer.'),
            comingSoon(
              'lasso',
              'Lasso Select',
              Target,
              'Lasso selection needs a workbook ink layer.',
            ),
            comingSoon('pen-black', 'Pen', PenLine, 'Ink drawing needs a workbook drawing layer.'),
            comingSoon(
              'pen-red',
              'Red Pen',
              PenLine,
              'Ink drawing needs a workbook drawing layer.',
            ),
            comingSoon(
              'highlighter',
              'Highlighter',
              Highlighter,
              'Ink drawing needs a workbook drawing layer.',
            ),
            command(
              'add-ink',
              'Add',
              Plus,
              'Describe the drawing or annotation you want to add to this workbook.',
            ),
            comingSoon(
              'ink-shape',
              'Ink to Shape',
              Target,
              'Ink-to-shape conversion needs a workbook drawing layer.',
            ),
            comingSoon(
              'ink-math',
              'Ink to Math',
              FunctionSquare,
              'Ink-to-math conversion needs a workbook drawing layer.',
            ),
            comingSoon(
              'ink-replay',
              'Ink Replay',
              RefreshCw,
              'Ink replay needs a workbook drawing layer.',
            ),
            command(
              'ink-help',
              'Ink Help',
              HelpCircle,
              'Explain how drawing and annotations work in this workspace.',
            ),
          ],
        },
      ];
    case 'page-layout':
      return [
        {
          label: 'Page Setup',
          actions: [
            command(
              'page-layout',
              'Page Layout',
              PanelTop,
              useSheet('Prepare a clean print-ready page layout'),
            ),
            command(
              'margins',
              'Margins',
              PanelTop,
              useSheet('Recommend page margins for printing'),
            ),
            command(
              'orientation',
              'Orientation',
              MoveHorizontal,
              useSheet('Choose portrait or landscape orientation for the data'),
            ),
            command(
              'print-area',
              'Print Area',
              Grid2X2,
              useSheet('Set the useful data range as the print area'),
            ),
            command(
              'page-breaks',
              'Breaks',
              Columns3,
              useSheet('Suggest page breaks for printing'),
            ),
          ],
        },
        {
          label: 'Scale & Sheet Options',
          actions: [
            command(
              'scale',
              'Scale to Fit',
              SlidersHorizontal,
              useSheet('Fit the data to a readable printed page'),
            ),
            command(
              'gridlines',
              'Gridlines',
              Grid2X2,
              useSheet('Explain or prepare gridline visibility for this sheet'),
            ),
            command(
              'headings',
              'Headings',
              Columns3,
              useSheet('Explain or prepare row and column headings for printing'),
            ),
            comingSoon(
              'themes',
              'Themes',
              Paintbrush,
              'Workbook theme preservation needs workbook metadata support.',
            ),
          ],
        },
      ];
    case 'formulas':
      return [
        {
          label: 'Function Library',
          actions: [
            command(
              'insert-function',
              'Insert Function',
              FunctionSquare,
              useSheet('Help me write and verify the right formula'),
            ),
            command(
              'autosum-formula',
              'AutoSum',
              Sigma,
              useSheet('Insert a verified sum formula for the relevant range'),
            ),
            command(
              'recent-functions',
              'Recently Used',
              FileClock,
              'Show the most useful formula patterns for this workbook.',
            ),
            command(
              'financial',
              'Financial',
              Calculator,
              useSheet('Explain or build a financial formula'),
            ),
            command('logical', 'Logical', Check, useSheet('Explain or build a logical formula')),
            command(
              'text-functions',
              'Text',
              TextCursorInput,
              useSheet('Explain or build a text formula'),
            ),
            command(
              'date-functions',
              'Date & Time',
              FileClock,
              useSheet('Explain or build a date and time formula'),
            ),
            command(
              'lookup',
              'Lookup & Reference',
              SearchCheck,
              useSheet('Explain or build a lookup formula'),
            ),
            command('math', 'Math & Trig', Sigma, useSheet('Explain or build a math formula')),
            command(
              'more-functions',
              'More Functions',
              MoreHorizontal,
              'List supported formulas and explain which one fits the request.',
            ),
          ],
        },
        {
          label: 'Defined Names',
          actions: [
            command(
              'name-manager',
              'Name Manager',
              Tags,
              'Inspect named-range needs and suggest safe names for this workbook.',
            ),
            command(
              'define-name',
              'Define Name',
              Tags,
              useSheet('Define a meaningful name for the selected range'),
            ),
            comingSoon(
              'use-formula',
              'Use in Formula',
              FunctionSquare,
              'Named-range persistence needs workbook metadata support.',
            ),
            command(
              'create-from-selection',
              'Create from Selection',
              Tags,
              useSheet('Create names from the selected headers'),
            ),
          ],
        },
        {
          label: 'Formula Auditing',
          actions: [
            command(
              'precedents',
              'Trace Precedents',
              SearchCheck,
              useSheet('Explain which cells feed the selected formula'),
            ),
            command(
              'dependents',
              'Trace Dependents',
              SearchCheck,
              useSheet('Explain which formulas depend on the selected cell'),
            ),
            command(
              'remove-arrows',
              'Remove Arrows',
              X,
              'Clear formula-audit annotations from the workspace.',
            ),
            command(
              'show-formulas',
              'Show Formulas',
              Code2,
              useSheet('Audit and explain all formulas in this sheet'),
            ),
            command(
              'error-checking',
              'Error Checking',
              ShieldCheck,
              useSheet('Find and explain formula errors'),
            ),
            command(
              'evaluate-formula',
              'Evaluate Formula',
              Calculator,
              useSheet('Evaluate the selected formula step by step'),
            ),
            comingSoon(
              'watch-window',
              'Watch Window',
              FileClock,
              'Persistent formula watches need a workbook analysis panel.',
            ),
          ],
        },
        {
          label: 'Calculation',
          actions: [
            command(
              'calculation-options',
              'Calculation Options',
              SlidersHorizontal,
              'Explain the calculation mode and formula refresh behavior.',
            ),
            command(
              'calculate-now',
              'Calculate Now',
              Calculator,
              'Recalculate and verify workbook formulas now.',
            ),
            command(
              'calculate-sheet',
              'Calculate Sheet',
              Sigma,
              useSheet('Recalculate and verify formulas'),
            ),
          ],
        },
      ];
    case 'data':
      return [
        {
          label: 'Get & Transform',
          actions: [
            command(
              'get-data',
              'Get Data',
              Database,
              'Help me import or connect data safely into this workbook.',
            ),
            command(
              'from-csv',
              'From Text/CSV',
              Upload,
              'Import a Text or CSV file into a new worksheet.',
            ),
            command(
              'from-picture',
              'From Picture',
              Image,
              'Extract a table from an image after showing the proposed data.',
            ),
            command(
              'from-web',
              'From Web',
              CloudIcon,
              'Find and import a bounded web data source after showing its source.',
            ),
            command(
              'from-table',
              'From Table/Range',
              Table2,
              useSheet('Turn this range into a clean data source'),
            ),
            command(
              'refresh-all',
              'Refresh All',
              RefreshCw,
              'Refresh available workbook data sources and explain any that are unavailable.',
            ),
            command(
              'recent-sources',
              'Recent Sources',
              FileClock,
              'Show recent local data sources used in this workspace.',
            ),
            comingSoon(
              'queries',
              'Queries & Connections',
              Database,
              'Persistent external connections need a server-side connector.',
            ),
            comingSoon(
              'existing-connections',
              'Existing Connections',
              Database,
              'Persistent external connections need a server-side connector.',
            ),
          ],
        },
        {
          label: 'Data Types',
          actions: [
            command(
              'stocks',
              'Stocks',
              BarChart3,
              useSheet('Identify or explain stock-like columns'),
            ),
            command(
              'currencies',
              'Currencies',
              Calculator,
              useSheet('Identify or explain currency columns'),
            ),
            command(
              'geography',
              'Geography',
              CloudIcon,
              useSheet('Identify or explain geography columns'),
            ),
          ],
        },
        {
          label: 'Sort & Filter',
          actions: [
            command('sort', 'Sort', SortAsc, useSheet('Sort this data using the selected column')),
            command(
              'filter',
              'Filter',
              Filter,
              useSheet('Filter this data using the selected criteria'),
            ),
            command(
              'advanced-filter',
              'Advanced',
              SlidersHorizontal,
              useSheet('Build an advanced filter for this data'),
            ),
          ],
        },
        {
          label: 'Data Tools',
          actions: [
            command(
              'text-columns',
              'Text to Columns',
              Columns3,
              useSheet('Split the selected text column into columns'),
            ),
            command(
              'flash-fill',
              'Flash Fill',
              WandSparkles,
              useSheet('Infer and fill the pattern from the selected examples'),
            ),
            command(
              'remove-duplicates',
              'Remove Duplicates',
              Eraser,
              useSheet('Find duplicate records and prepare a safe removal preview'),
            ),
            command(
              'validation',
              'Data Validation',
              ListChecksIcon,
              useSheet('Recommend validation rules for the selected field'),
            ),
            command(
              'consolidate',
              'Consolidate',
              Database,
              'Consolidate compatible sheets into one verified sheet.',
            ),
            comingSoon(
              'data-model',
              'Data Model',
              Database,
              'Persistent relational data models need a workbook model layer.',
            ),
            command(
              'analyze-data',
              'Analyze Data',
              BarChart3,
              'Run a grounded analyst briefing for this workbook.',
            ),
          ],
        },
        {
          label: 'Forecast & Outline',
          actions: [
            command(
              'what-if',
              'What-if Analysis',
              Target,
              useSheet('Run a what-if analysis with explicit assumptions'),
            ),
            command(
              'forecast',
              'Forecast Sheet',
              BarChart3,
              useSheet('Prepare a forecast using the available numeric time series'),
            ),
            command('group', 'Group', Columns3, useSheet('Group the selected rows or columns')),
            command(
              'ungroup',
              'Ungroup',
              Columns3,
              useSheet('Remove grouping from the selected rows or columns'),
            ),
            command('subtotal', 'Subtotal', Sigma, useSheet('Add verified subtotals to this data')),
          ],
        },
      ];
    case 'review':
      return [
        {
          label: 'Proofing',
          actions: [
            command(
              'spelling',
              'Spelling',
              TextSelect,
              'Check workbook text for spelling issues and propose corrections.',
            ),
            command(
              'thesaurus',
              'Thesaurus',
              TextSelect,
              'Suggest clearer alternatives for the selected text.',
            ),
            command(
              'statistics',
              'Workbook Statistics',
              BarChart3,
              'Show grounded workbook statistics and quality metrics.',
            ),
            command(
              'performance',
              'Check Performance',
              GaugeIcon,
              'Check this workbook for performance risks and explain findings.',
            ),
            command(
              'accessibility',
              'Check Accessibility',
              ShieldCheck,
              'Audit this workbook for accessibility issues.',
            ),
            command(
              'translate',
              'Translate',
              TextSelect,
              'Translate the selected text while preserving the workbook structure.',
            ),
            command(
              'changes',
              'Show Changes',
              FileClock,
              'Show the verified history of changes in this session.',
            ),
          ],
        },
        {
          label: 'Comments & Notes',
          actions: [
            command(
              'new-comment',
              'New Comment',
              MessageSquare,
              'Add a comment to the selected cell.',
            ),
            command(
              'delete-comment',
              'Delete',
              X,
              'Remove the selected cell comment after previewing the change.',
            ),
            command(
              'previous-comment',
              'Previous',
              MessageSquare,
              'Go to the previous comment in this workbook.',
            ),
            command(
              'next-comment',
              'Next',
              MessageSquare,
              'Go to the next comment in this workbook.',
            ),
            command(
              'show-comments',
              'Show Comments',
              MessageSquare,
              'Summarize comments and notes in this workbook.',
            ),
            command('notes', 'Notes', FileClock, 'Show or create notes for the selected cells.'),
          ],
        },
        {
          label: 'Protect',
          actions: [
            command(
              'protect-sheet',
              'Protect Sheet',
              LockKeyhole,
              'Explain and prepare sheet protection settings.',
            ),
            command(
              'protect-workbook',
              'Protect Workbook',
              LockKeyhole,
              'Explain and prepare workbook protection settings.',
            ),
            command(
              'allow-edit',
              'Allow Edit Ranges',
              ShieldCheck,
              'Define which ranges should remain editable.',
            ),
            command(
              'unshare',
              'Unshare Workbook',
              X,
              'Explain sharing status and how to remove sharing safely.',
            ),
          ],
        },
        {
          label: 'Ink',
          actions: [
            command(
              'hide-ink',
              'Hide Ink',
              PenLine,
              'Hide or explain ink annotations in this workspace.',
            ),
          ],
        },
      ];
    case 'automate':
      return [
        {
          label: 'Automate',
          actions: [
            command(
              'automate-task',
              'Automate a Task',
              WandSparkles,
              'Understand my request, inspect the workbook, and prepare the safest multi-step automation plan.',
            ),
            command(
              'record-workflow',
              'Record Workflow',
              FileClock,
              'Turn my next reviewed actions into a reusable workflow.',
            ),
            command(
              'run-workflow',
              'Run Workflow',
              RefreshCw,
              'Show available verified workflows and run the one I choose.',
            ),
            command(
              'agent-tools',
              'Agent Tools',
              Code2,
              'Show every workbook tool available to the agent and what each one can do.',
            ),
          ],
        },
        {
          label: 'Connected Services',
          actions: [
            command(
              'connection-health',
              'Connection Health',
              ShieldCheck,
              'Check model and workspace connection health.',
            ),
            command(
              'send-feedback',
              'Send Feedback',
              MessageSquare,
              'Help me report a problem or suggest a workspace improvement.',
            ),
          ],
        },
      ];
    case 'help':
      return [
        {
          label: 'Help',
          actions: [
            command(
              'help',
              'Help',
              HelpCircle,
              'Explain how to use this workspace in simple language.',
            ),
            command(
              'search-help',
              'Search Help',
              SearchCheck,
              'Search the workspace help for my question.',
            ),
            command(
              'training',
              'Training',
              Presentation,
              'Teach me how to complete my task step by step.',
            ),
            command(
              'keyboard',
              'Keyboard Shortcuts',
              Code2,
              'Show keyboard shortcuts for this workspace.',
            ),
            command(
              'accessibility-help',
              'Accessibility',
              ShieldCheck,
              'Explain accessibility features and help me use them.',
            ),
          ],
        },
        {
          label: 'Workspace',
          actions: [
            command(
              'workspace-integrity',
              'Verify Integrity',
              ShieldCheck,
              'Run a complete read-only integrity check of this workbook and report actionable findings.',
            ),
            command('export', 'Export', Download, 'Export this verified workbook to .xlsx.'),
            command(
              'open-tools',
              'All Tools',
              Code2,
              'Show all workspace and agent tools in one place.',
            ),
          ],
        },
      ];
  }
}

function GaugeIcon({ size = 14 }: { size?: number }) {
  return (
    <span aria-hidden="true" style={{ fontSize: size, lineHeight: 1 }}>
      ◔
    </span>
  );
}

function CloudIcon({ size = 14, strokeWidth = 1.8 }: { size?: number; strokeWidth?: number }) {
  return (
    <span
      aria-hidden="true"
      style={{ fontSize: size, lineHeight: 1, fontWeight: strokeWidth > 1 ? 600 : 400 }}
    >
      ☁
    </span>
  );
}

function ListChecksIcon({ size = 14 }: { size?: number }) {
  return (
    <span aria-hidden="true" style={{ fontSize: size, lineHeight: 1 }}>
      ☷
    </span>
  );
}

function FlagIcon({ size = 14 }: { size?: number }) {
  return (
    <span aria-hidden="true" style={{ fontSize: size, lineHeight: 1 }}>
      ⚑
    </span>
  );
}

export const WorkspaceRibbon: React.FC<WorkspaceRibbonProps> = ({
  activeSheetName,
  isProcessing,
  onManualAction,
  onOpenOperationModal,
  onUploadFile,
  onExport,
  onUndo,
  onRedo,
  canUndo,
  canRedo,
}) => {
  const [activeTab, setActiveTab] = useState<RibbonTabId>('home');
  const [showAll, setShowAll] = useState(false);
  const groups = useMemo(
    () => actionGroups(activeTab, activeSheetName),
    [activeTab, activeSheetName],
  );

  const runAction = (action: RibbonAction) => {
    if (action.disabled || isProcessing) return;
    if (action.id === 'draw-undo') return onUndo();
    if (action.id === 'draw-redo') return onRedo();
    if (action.id === 'export') return onExport();
    if (action.id === 'get-data' || action.id === 'from-csv') return onUploadFileRequest();
    if (action.id === 'find-select' || action.id === 'search-help')
      return onManualAction(action.id);
    onManualAction(action.id);
  };

  const uploadInput = useRef<HTMLInputElement>(null);
  const onUploadFileRequest = () => uploadInput.current?.click();

  const renderTab = (tab: { id: RibbonTabId; label: string }, includeHeading: boolean) => (
    <React.Fragment key={tab.id}>
      {includeHeading && <div className="workspace-ribbon-all-tab-label">{tab.label}</div>}
      <div className="workspace-ribbon-groups">
        {(includeHeading ? actionGroups(tab.id, activeSheetName) : groups).map((group) => (
          <div className="workspace-ribbon-group" key={`${tab.id}-${group.label}`}>
            <div className="workspace-ribbon-actions">
              {group.actions.map((action) => {
                const Icon = action.icon;
                return (
                  <button
                    key={action.id}
                    type="button"
                    className={`workspace-ribbon-action is-${action.tone ?? 'default'}`}
                    disabled={
                      Boolean(action.disabled) ||
                      isProcessing ||
                      (action.id === 'draw-undo' && !canUndo) ||
                      (action.id === 'draw-redo' && !canRedo)
                    }
                    title={
                      action.hint ??
                      (action.disabled
                        ? 'Visible here; this capability needs richer workbook fidelity.'
                        : `Use ${action.label} manually on the selected cells`)
                    }
                    onClick={() => runAction(action)}
                  >
                    <Icon size={15} strokeWidth={1.8} />
                    <span>{action.label}</span>
                  </button>
                );
              })}
            </div>
            <span className="workspace-ribbon-group-label">{group.label}</span>
          </div>
        ))}
      </div>
    </React.Fragment>
  );

  return (
    <section className="workspace-ribbon" aria-label="Workspace tools">
      <div className="workspace-ribbon-tabs" role="tablist" aria-label="Workbook commands">
        {tabs.map((tab) => (
          <button
            key={tab.id}
            type="button"
            role="tab"
            aria-selected={activeTab === tab.id}
            className={`workspace-ribbon-tab ${activeTab === tab.id ? 'active' : ''}`}
            onClick={() => {
              setActiveTab(tab.id);
              setShowAll(false);
            }}
          >
            {tab.label}
          </button>
        ))}
        <button
          type="button"
          className="workspace-ribbon-more"
          aria-label="Show all workspace tools"
          title="Show all workspace tools"
          onClick={() => setShowAll((value) => !value)}
        >
          <MoreHorizontal size={15} />
        </button>
      </div>
      <div
        className="workspace-ribbon-toolbar"
        role="tabpanel"
        aria-label={`${activeTab} commands`}
      >
        <div className="workspace-ribbon-quick-actions" aria-label="Quick workspace actions">
          <button
            type="button"
            className="workspace-ribbon-quick"
            onClick={onUploadFileRequest}
            disabled={isProcessing}
            title="Upload workbook or CSV"
          >
            <FolderOpen size={15} /> Open
          </button>
          <button
            type="button"
            className="workspace-ribbon-quick"
            onClick={onExport}
            disabled={isProcessing}
            title="Export workbook"
          >
            <Download size={15} /> Export
          </button>
          <button
            type="button"
            className="workspace-ribbon-quick"
            onClick={onOpenOperationModal}
            disabled={isProcessing}
            title="Open deterministic engine operations"
          >
            <SlidersHorizontal size={15} /> Operations
          </button>
        </div>
        <input
          ref={uploadInput}
          className="workspace-ribbon-file-input"
          type="file"
          accept=".xlsx,.xls,.csv"
          tabIndex={-1}
          aria-hidden="true"
          onChange={(event) => {
            const file = event.target.files?.[0];
            if (file) onUploadFile(file);
            event.target.value = '';
          }}
        />
        {showAll
          ? tabs.map((tab) => renderTab(tab, true))
          : renderTab({ id: activeTab, label: activeTab }, false)}
      </div>
      <div className="workspace-ribbon-status" role="status" aria-live="polite">
        <span>
          <FileSpreadsheet size={13} /> Active: {activeSheetName}
        </span>
        <span>
          <ShieldCheck size={13} /> Guarded agent tools
        </span>
        {isProcessing && (
          <span className="workspace-ribbon-busy">
            <RefreshCw size={12} /> Working…
          </span>
        )}
      </div>
    </section>
  );
};
