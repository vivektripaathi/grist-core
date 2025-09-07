/**
 * Interface for column information
 */
export interface IColumnInfo {
  id: string;
  label: string;
  type: string;
}

/**
 * Interface for bulk data
 */
export interface IBulkData {
  [columnId: string]: any[];
}

/**
 * Interface for sheet creation options
 */
export interface ICreateSheetOptions {
  baseTableName: string;
  columns: IColumnInfo[];
  bulkData: IBulkData;
  navigateToSheet?: boolean;
}

/**
 * Utility class for creating new sheets from various data sources
 * Provides reusable functionality to reduce code redundancy
 */
export class SheetCreationUtils {

  /**
   * Creates a new sheet with the provided data
   */
  public static async createNewSheet(
    gristDoc: any,
    options: ICreateSheetOptions
  ): Promise<string> {
    const {
      baseTableName,
      columns,
      bulkData,
      navigateToSheet = true
    } = options;

    // Generate unique table name
    const newTableId = this._generateUniqueTableName(gristDoc, baseTableName);

    // Add the new table
    await gristDoc.docModel.docData.sendAction([
      'AddTable',
      newTableId,
      columns
    ]);

    // Add the data if provided
    if (Object.keys(bulkData).length > 0) {
      // Calculate the number of rows from the first column's data
      const firstColumnData = Object.values(bulkData)[0];
      const numRows = firstColumnData ? firstColumnData.length : 0;
      
      // Create array of nulls for new row IDs
      const rowIds = new Array(numRows).fill(null);
      
      await gristDoc.docModel.docData.sendAction([
        'BulkAddRecord',
        newTableId,
        rowIds,
        bulkData
      ]);
    }

    // Navigate to the new sheet if requested
    if (navigateToSheet) {
      await this._navigateToSheet(gristDoc, newTableId);
    }

    return newTableId;
  }

  /**
   * Converts BaseView selection data to the format needed for sheet creation
   */
  public static convertBaseViewSelection(selection: any): {columns: IColumnInfo[], bulkData: IBulkData} {
    const columns: IColumnInfo[] = [];
    const bulkData: IBulkData = {};

    // Prepare columns based on selected fields
    for (const field of selection.fields) {
      columns.push({
        label: field.label(),
        id: field.colId(),
        type: field.column().type()
      });
    }

    // Prepare bulk data
    const colIds = columns.map(col => col.id);
    colIds.forEach(colId => {
      bulkData[colId] = [];
    });

    // Fill data for each row
    for (const rowId of selection.rowIds) {
      for (let i = 0; i < selection.fields.length; i++) {
        const field = selection.fields[i];
        const colId = colIds[i];
        const cellValue = field.column().rawData().getValue(rowId);
        bulkData[colId].push(cellValue);
      }
    }

    return {columns, bulkData};
  }

  /**
   * Converts global selection data to the format needed for sheet creation
   * Arranges selections side by side horizontally
   */
  public static convertGlobalSelectionData(formattedData: any[]): {columns: IColumnInfo[], bulkData: IBulkData} {
    if (!formattedData || formattedData.length === 0) {
      throw new Error('No formatted data provided for sheet creation');
    }

    // Calculate dimensions - arrange selections side by side
    const maxRows = Math.max(...formattedData.map(entry => entry.rowCount || 0));
    const totalColumns = formattedData.reduce((sum, entry) => sum + (entry.columnCount || 0), 0);

    console.log(`Arranging ${formattedData.length} selections side by side:`,
                `${maxRows} rows × ${totalColumns} columns`);

    const columns: IColumnInfo[] = [];
    const bulkData: IBulkData = {};
    const usedColumnIds = new Set<string>();

    // Process each selection side by side
    for (let selectionIndex = 0; selectionIndex < formattedData.length; selectionIndex++) {
      const entry = formattedData[selectionIndex];
      if (!entry.headers || !Array.isArray(entry.headers)) {
        console.warn(`Selection ${selectionIndex + 1} has invalid headers, skipping`);
        continue;
      }

      // Use table name as prefix instead of generic S1, S2, etc.
      const tableName = entry.table || `S${selectionIndex + 1}`;
      // Sanitize table name for use as column prefix (remove spaces, special chars)
      const safeTableName = tableName.replace(/[^a-zA-Z0-9_]/g, '_');
      const prefix = `${safeTableName}_`;

      // Add columns for this selection based on headers
      for (const header of entry.headers) {
        let columnId = `${prefix}${header}`;
        let columnLabel = `${prefix}${header}`;
        
        // Handle duplicate column IDs by adding a counter
        let counter = 1;
        const baseColumnId = columnId;
        while (usedColumnIds.has(columnId)) {
          columnId = `${baseColumnId}_${counter}`;
          columnLabel = `${prefix}${header}_${counter}`;
          counter++;
        }
        usedColumnIds.add(columnId);

        columns.push({
          id: columnId,
          label: columnLabel,
          type: 'Text'
        });

        // Initialize column data array
        bulkData[columnId] = [];

        // Fill data for this column, padding with empty strings if needed
        for (let rowIndex = 0; rowIndex < maxRows; rowIndex++) {
          let value = '';
          if (entry.data && Array.isArray(entry.data) && rowIndex < entry.data.length) {
            value = entry.data[rowIndex][header] || '';
          }
          bulkData[columnId].push(value);
        }
      }
    }

    if (columns.length === 0) {
      throw new Error('No valid columns found in formatted data');
    }

    return {columns, bulkData};
  }

  /**
   * Generates a unique table name with retry logic
   */
  private static _generateUniqueTableName(
    gristDoc: any,
    baseTableName: string,
    maxRetries: number = 10
  ): string {
    let counter = 1;
    let tableName = baseTableName;

    while (counter <= maxRetries) {
      if (!gristDoc.docModel.tables.rowModels.find((t: any) => t.tableId.peek() === tableName)) {
        return tableName;
      }
      tableName = `${baseTableName}_${counter}`;
      counter++;
    }

    throw new Error(`Could not generate unique table name after ${maxRetries} attempts`);
  }

  private static async _navigateToSheet(gristDoc: any, tableId: string): Promise<void> {
    const newTableRec = gristDoc.docModel.tables.rowModels.find((t: any) =>
      t.tableId.peek() === tableId);

    if (newTableRec) {
      const newViewSection = newTableRec.rawViewSectionRef.peek();
      if (newViewSection) {
        await gristDoc.openDocPage(newViewSection);
      } else {
        console.error('No view section found for new table');
      }
    } else {
      console.error('New table not found');
    }
  }
}
