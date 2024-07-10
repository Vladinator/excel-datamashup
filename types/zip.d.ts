import type { Buffer } from 'buffer/';

export type UnzippedItem = {
    path: string;
    type: 'File' | 'Directory';
    size: number;
    data: Buffer | string;
};
