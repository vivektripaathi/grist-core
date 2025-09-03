/**
 * Client-side validation utilities for file imports.
 * Provides immediate feedback for invalid filenames and Excel content before upload.
 */

import * as path from 'path';
import ExcelJS from 'exceljs';

export interface ImportFileValidationOptions {
  allowSpecialCharsInFilename?: boolean;
  allowSpacesInFilename?: boolean;
  allowSpecialCharsInSheetName?: boolean;
  allowSpacesInSheetName?: boolean;
  allowMergedCells?: boolean;
}


export interface ImportFileValidationResult {
  isValid: boolean;
  errors: string[];
  warnings: string[];
}

/**
 * Validates text for special characters and spaces
 */
function validateNameText(
  text: string,
  type: 'filename' | 'sheet name',
  options: { allowSpecialChars?: boolean; allowSpaces?: boolean }
): string[] {
  const errors: string[] = [];

  // Check for special characters
  const specialCharsRegex = /[!@#$%^&*()+=[\]{}|\\:";'<>?,]/;
  if (!options.allowSpecialChars && specialCharsRegex.test(text)) {
    errors.push(
      `${type.charAt(0).toUpperCase() + type.slice(1)} "${text}" contains special characters. `
    );
  }

  // Check for spaces
  const spaceRegex = /\s/;
  if (!options.allowSpaces && spaceRegex.test(text)) {
    errors.push(
      `${type.charAt(0).toUpperCase() + type.slice(1)} "${text}" contains spaces.`
    );
  }

  return errors;
}


/**
 * Validates filename on the client side before upload
 */
export function validateFilename(
  filename: string,
  options: ImportFileValidationOptions = {}
): ImportFileValidationResult {
  console.log(`Validating file name: ${filename}`);
  const result: ImportFileValidationResult = {
    isValid: true,
    errors: [],
    warnings: []
  };

  const baseName = path.basename(filename, path.extname(filename));

  const errors = validateNameText(baseName, 'filename', {
    allowSpecialChars: options.allowSpecialCharsInFilename,
    allowSpaces: options.allowSpacesInFilename
  });

  result.errors.push(...errors);
  result.isValid = result.errors.length === 0;
  return result;
}


/**
 * Checks if a file is an Excel file based on its extension
 */
function isExcelFile(filename: string): boolean {
  const ext = path.extname(filename).toLowerCase();
  return ['.xlsx', '.xls', '.xlsm'].includes(ext);
}


/**
 * Validates Excel file content including sheet names and merged cells
 */
async function validateExcelFile(
  file: File,
  options: ImportFileValidationOptions = {}
): Promise<ImportFileValidationResult> {
  const result: ImportFileValidationResult = {
    isValid: true,
    errors: [],
    warnings: []
  };

  try {
    const workbook = new ExcelJS.Workbook();
    const arrayBuffer = await file.arrayBuffer();

    await workbook.xlsx.load(arrayBuffer);

    workbook.eachSheet((worksheet, sheetId) => {
      const sheetName = worksheet.name;

      const sheetNameErrors = validateNameText(sheetName, 'sheet name', {
        allowSpecialChars: options.allowSpecialCharsInSheetName,
        allowSpaces: options.allowSpacesInSheetName
      });
      result.errors.push(...sheetNameErrors);

      if (!options.allowMergedCells && worksheet.hasMerges) {
        result.errors.push(
          `Sheet "${sheetName}": Contains merged cells. `
        );
      }
    });

  } catch (error) {
    result.warnings.push(
      `Could not validate Excel file content: ${error instanceof Error ? error.message : 'Unknown error'}. ` +
      `Server-side validation will be performed.`
    );
  }

  result.isValid = result.errors.length === 0;
  return result;
}


/**
 * Validates all files in a FileList for import
 */
export async function validateFilesForImport(
  files: File[],
  options: ImportFileValidationOptions = {}
): Promise<ImportFileValidationResult> {
  const result: ImportFileValidationResult = {
    isValid: true,
    errors: [],
    warnings: []
  };

  for (const file of files) {
    const filenameValidation = validateFilename(file.name, options);
    result.errors.push(...filenameValidation.errors);
    result.warnings.push(...filenameValidation.warnings);

    if (isExcelFile(file.name)) {
      const excelValidation = await validateExcelFile(file, options);
      result.errors.push(...excelValidation.errors);
      result.warnings.push(...excelValidation.warnings);
    }
  }

  result.isValid = result.errors.length === 0;
  return result;
}