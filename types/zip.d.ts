import type { Buffer } from 'buffer/';
import type { ParseResult, ParseError } from './index.d';

export type UnzippedItem = {
    path: string;
    type: 'File' | 'Directory';
    size: number;
    data: Buffer | string;
};

export type UnzippedExcelDataMashup = {
    file: UnzippedItem;
    xml: string;
} & (
    | { error?: never; result: ParseResult }
    | { error: ParseError; result?: never }
);

export type UnzippedExcel = {
    files: UnzippedItem[];
    datamashup?: UnzippedExcelDataMashup;
    getFormula: () => string | undefined;
    setFormula: (formula: string) => void;
    save: () => Promise<Buffer>;
};
