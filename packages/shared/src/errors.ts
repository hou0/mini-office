export class OfficeError extends Error {
  readonly code: string;
  readonly detail?: string;

  constructor(code: string, message: string, detail?: string) {
    super(message);
    this.name = 'OfficeError';
    this.code = code;
    this.detail = detail;
  }
}

export const ErrorCodes = {
  UNSUPPORTED_FORMAT: 'UNSUPPORTED_FORMAT',
  CORRUPT_PACKAGE: 'CORRUPT_PACKAGE',
  MACRO_DOCUMENT: 'MACRO_DOCUMENT',
  LEGACY_FORMAT: 'LEGACY_FORMAT',
  PASSWORD_PROTECTED: 'PASSWORD_PROTECTED',
  PARSE_FAILED: 'PARSE_FAILED',
  SAVE_FAILED: 'SAVE_FAILED',
  PERMISSION_DENIED: 'PERMISSION_DENIED',
  EXPORT_FAILED: 'EXPORT_FAILED',
  CANCELLED: 'CANCELLED',
  TOO_LARGE: 'TOO_LARGE',
} as const;

export function toOfficeError(error: unknown, fallbackCode = ErrorCodes.PARSE_FAILED): OfficeError {
  if (error instanceof OfficeError) return error;
  if (error instanceof Error) return new OfficeError(fallbackCode, error.message, error.stack);
  return new OfficeError(fallbackCode, String(error));
}
