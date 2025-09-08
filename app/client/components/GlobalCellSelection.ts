import {CopySelection} from 'app/client/components/CopySelection';
import {ViewFieldRec} from 'app/client/models/entities/ViewFieldRec';
import {UIRowId} from 'app/plugin/GristAPI';
import {Computed, Disposable, Observable} from 'grainjs';
import {SheetCreationUtils} from 'app/client/lib/SheetCreationUtils';
import {showNewSheetNameModal} from 'app/client/ui/NewSheetNameModal';

/**
 * Interface for a single global cell selection entry
 */
export interface IGlobalSelectionEntry {
  /** Unique identifier for this selection */
  id: string;
  /** Name of the sheet/view where this selection was made */
  sheetName: string;
  /** ID of the view section */
  sectionId: number;
  /** Table ID */
  tableId: string;
  /** Selected row IDs */
  rowIds: UIRowId[];
  /** Selected field records */
  fields: ViewFieldRec[];
  /** Timestamp when selection was added */
  timestamp: number;
  /** Display data for the cells (values with headers) */
  cellData: {[rowId: string]: {[colId: string]: any}};
  /** Column headers */
  headers: string[];
}

/**
 * GlobalCellSelection manages global cell selections across multiple sheets.
 * This allows users to select cells from different sheets and collect them for analysis.
 */
export class GlobalCellSelection extends Disposable {
  /** Observable array of global selections */
  public readonly selections: Observable<IGlobalSelectionEntry[]>;
  
  /** Computed observable to check if there are any selections */
  public readonly hasSelections: Computed<boolean>;

  constructor() {
    super();
    this.selections = Observable.create(this, []);
    this.hasSelections = Computed.create(this, this.selections, (use, selections) => selections.length > 0);
  }

  /**
   * Add a selection to the global collection
   */
  public addSelection(
    sheetName: string,
    sectionId: number,
    tableId: string,
    copySelection: CopySelection
  ): void {
    const id = `${sectionId}-${Date.now()}-${Math.random().toString(36).substr(2, 9)}`;

    // Extract cell data with headers
    const cellData: {[rowId: string]: {[colId: string]: any}} = {};
    const headers: string[] = [];

    // Get headers from fields
    copySelection.fields.forEach(field => {
      headers.push(field.label.peek());
    });

    // Get cell values using the CopySelection's columns data
    copySelection.rowIds.forEach(rowId => {
      cellData[rowId] = {};
      copySelection.columns.forEach((column, index) => {
        const field = copySelection.fields[index];
        const colId = field.colId.peek();
        const formattedValue = column.fmtGetter(rowId);
        cellData[rowId][colId] = formattedValue;
      });
    });

    const entry: IGlobalSelectionEntry = {
      id,
      sheetName,
      sectionId,
      tableId,
      rowIds: [...copySelection.rowIds],
      fields: [...copySelection.fields],
      timestamp: Date.now(),
      cellData,
      headers
    };

    const currentSelections = this.selections.get();
    this.selections.set([...currentSelections, entry]);
  }

  /**
   * Clear all global selections
   */
  public clearSelections(): void {
    this.selections.set([]);
  }

  /**
   * Remove a specific selection by ID
   */
  public removeSelection(id: string): void {
    const currentSelections = this.selections.get();
    this.selections.set(currentSelections.filter(entry => entry.id !== id));
  }

  /**
   * Get formatted data for console display
   */
  public getFormattedData(): any[] {
    const selections = this.selections.get();

    return selections.filter(entry => entry != null).map(entry => {
      // Add validation to ensure entry has required properties
      if (!entry) {
        console.warn('Encountered null or undefined selection entry, skipping');
        return null;
      }

      if (!entry.tableId || !entry.sheetName || !entry.rowIds || !entry.fields) {
        console.warn('Selection entry missing required properties:', entry);
        return null;
      }

      const formattedEntry = {
        id: entry.id,
        sheet: entry.sheetName,
        table: entry.tableId,
        timestamp: new Date(entry.timestamp).toLocaleString(),
        rowCount: entry.rowIds.length,
        columnCount: entry.fields.length,
        headers: entry.headers,
        data: [] as any[]
      };

      // Format data as rows with header mapping
      entry.rowIds.forEach(rowId => {
        const row: any = {rowId};
        entry.fields.forEach(field => {
          try {
            const colId = field.column.peek().colId.peek();
            const header = field.label.peek();
            row[header] = entry.cellData[rowId]?.[colId] || null;
          } catch (error) {
            console.warn('Error processing field data:', error, field);
            // Continue with other fields
          }
        });
        formattedEntry.data.push(row);
      });

      return formattedEntry;
    }).filter(entry => entry != null); // Remove any null entries
  }

  /**
   * Creates a new sheet with all the global selection data using the reusable SheetCreationUtils
   * Arranges selections side by side horizontally instead of stacking vertically
   */
  public async createNewSheet(gristDoc: any): Promise<void> {
    const formattedData = this.getFormattedData();

    if (formattedData.length === 0) {
      console.log('No global selections available to create sheet');
      return;
    }

    showNewSheetNameModal(async (sheetName) => {
      try {
        // Convert global selection data to the format needed by SheetCreationUtils
        const {columns, bulkData} = SheetCreationUtils.convertGlobalSelectionData(formattedData);
        // Create new sheet using the reusable utility
        await SheetCreationUtils.createNewSheet(gristDoc, {
          baseTableName: sheetName,
          columns,
          bulkData,
          navigateToSheet: true
        });

        const maxRows = Math.max(...formattedData.map(entry => entry.rowCount));
        const totalColumns = columns.length;
        console.log(`✅ Created new sheet with ${maxRows} rows and ${totalColumns} columns ` +
                    `from ${formattedData.length} selections arranged side by side`);
        // Clear selections after successful sheet creation
        this.clearSelections();
      } catch (error) {
        console.error('❌ Failed to create new sheet:', error);
        throw error;
      }
    }, 'GlobalSelections', gristDoc);
  }
}
