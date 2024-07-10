import type { UnzippedItem } from './zip';

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
    getFormula: () => string;
    resetPermissions: () => void;
    save: () => Promise<string>;
};
