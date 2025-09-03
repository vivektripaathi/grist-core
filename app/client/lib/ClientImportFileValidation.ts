/**
 * Client-side validation utilities for file imports.
 * Provides immediate feedback for invalid filenames before upload.
 */

import * as path from 'path';

export interface ClientValidationOptions {
  allowSpecialCharsInFilename?: boolean;
  allowSpacesInFilename?: boolean;
}

export interface ClientValidationResult {
  isValid: boolean;
  errors: string[];
}

/**
 * Validates filename on the client side before upload
 */
export function validateFilename(
  filename: string,
  options: ClientValidationOptions = {}
): ClientValidationResult {
  const result: ClientValidationResult = {
    isValid: true,
    errors: []
  };

  const baseName = path.basename(filename, path.extname(filename));

  // Check for special characters
  const specialCharsRegex = /[!@#$%^&*()+=[\]{}|\\:";'<>?,]/;
  if (!options.allowSpecialCharsInFilename && specialCharsRegex.test(baseName)) {
    result.errors.push(
      'Filename contains special characters. Only letters, numbers, hyphens, and underscores are allowed.'
    );
  }

  // Check for spaces
  const spaceRegex = /\s/;
  if (!options.allowSpacesInFilename && spaceRegex.test(baseName)) {
    result.errors.push(
      'Filename contains spaces. Please use underscores or hyphens instead of spaces.'
    );
  }

  result.isValid = result.errors.length === 0;
  return result;
}

/**
 * Validates all files in a FileList for import
 */
export function validateFilesForImport(
  files: File[],
  options: ClientValidationOptions = {}
): ClientValidationResult {
  const result: ClientValidationResult = {
    isValid: true,
    errors: []
  };

  for (const file of files) {
    const fileValidation = validateFilename(file.name, options);
    if (!fileValidation.isValid) {
      result.errors.push(`File "${file.name}": ${fileValidation.errors.join(', ')}`);
    }
  }

  result.isValid = result.errors.length === 0;
  return result;
}
