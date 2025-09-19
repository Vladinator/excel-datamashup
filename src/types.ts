import type { Buffer } from 'buffer/';

export type ParseError =
    | 'DataMashupNotFound'
    | 'Base64DecodeError'
    | 'ParseRootError';

export type Metadata = {
    version: number;
    metadata: string;
    content: UnzippedItem[];
};

export type ParseResult = {
    version: number;
    packageParts: UnzippedItem[];
    permissions: string;
    metadata: Metadata;
    permissionBindings: number[];
    setFormula: (formula: string) => void;
    getFormula: () => string | undefined;
    resetPermissions: () => void;
    save: () => Promise<string>;
};

export type UnzippedItem<T = Buffer> = {
    path: string;
    type: 'File' | 'Directory';
    size: number;
    data: T | string;
};

export type UnzippedExcelDataMashup<T = Buffer> = {
    file: UnzippedItem<T>;
    xml: string;
} & (
    | { error?: never; result: ParseResult }
    | { error: ParseError; result?: never }
);

export type UnzippedExcel<T = Buffer> = {
    files: UnzippedItem<T>[];
    datamashup?: UnzippedExcelDataMashup<T>;
    getFormula: () => string | undefined;
    setFormula: (formula: string) => void;
    save: () => Promise<T>;
};
