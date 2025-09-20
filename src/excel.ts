import { ParserMetadata, ParserRoot } from './datamashup';
import type { Result } from './types';
import { type Uint8ArrayUtilsFrom, Uint8ArrayUtils } from './utils';
import { type UnzippedItem, Unzip, Zip } from './zip';

/** Matching RegExp to extract the DataMashup XML tag and the base64 binary data. */
const MashupBinaryRegExp = /<DataMashup[^>]*>(.*?)<\/DataMashup>/s;

/** The string that signifies the DataMashup XML tag. */
const MashupBinaryPrefix = '<DataMashup ';

/**
 * The top-level binary stream has package parts, which is another ZIP archive with specific files.
 *
 * One of these is `Section1.m` which is a `Power Query Formula` following some strict rules.
 *
 * References:
 *
 * https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-qdeff/a4c2d0b9-9a9d-452d-8802-d68339374d57
 *
 * https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-qdeff/31e4aedb-e1ae-4ade-948c-1d377184fd52
 */
const MashupFormulaSectionDefault = 'Section1.m';

/**
 * If the top-level binary stream permission bindings become cryptographically invalid, then
 * we need to reset the permissions XML to this default to indicate that content has changed outside of Excel.
 *
 * References:
 *
 * https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-qdeff/d0959ba8-ac8d-4bee-bb58-9a869d7b226a
 */
const MashupPermissionDefaults = `<?xml version="1.0" encoding="utf-8"?>\r\n<PermissionList xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">\r\n\t<CanEvaluateFuturePackages>false</CanEvaluateFuturePackages>\r\n\t<FirewallEnabled>true</FirewallEnabled>\r\n\t<WorkbookGroupType xsi:nil="true" />\r\n</PermissionList>`;

export type ExcelCustomXmlResult =
    | 'NotParsed'
    | 'DataMashupNotFound'
    | 'Base64DecodeError'
    | 'ParseRootError'
    | 'MetadataError'
    | 'Working'
    | 'PackageUnzipError'
    | 'MetadataUnzipError'
    | 'Success';

export type ExcelCustomXmlRootData = ReturnType<(typeof ParserRoot)['parse']>;

export type ExcelCustomXmlMetadata = ReturnType<
    (typeof ParserMetadata)['parse']
>;

export class ExcelCustomXml {
    private _xmlContent: string;
    private _parseResult: ExcelCustomXmlResult | undefined;
    private _mashupBase64: string | null | undefined;
    private _packageZip: ExcelZip | undefined;
    private _rootData: ExcelCustomXmlRootData | undefined;
    private _metaData: ExcelCustomXmlMetadata | undefined;
    private _metaDataZip: ExcelZip | undefined;
    private _metaDataXml:
        | ReturnType<typeof Uint8ArrayUtils.toStringEncoding>
        | undefined;

    public get parseResult(): ExcelCustomXmlResult {
        return this._parseResult || 'NotParsed';
    }

    public get mashupBase64(): string | undefined {
        if (this._mashupBase64 !== undefined) {
            return this._mashupBase64 || undefined;
        }
        const match = this._xmlContent.match(MashupBinaryRegExp);
        if (!match) {
            this._mashupBase64 = null;
            return undefined;
        }
        this._mashupBase64 = match[1].trim();
        return this._mashupBase64;
    }

    public set mashupBase64(value: string) {
        if (!value) {
            return;
        }
        const current = this.mashupBase64;
        if (!current) {
            return;
        }
        this._mashupBase64 = value;
        this._xmlContent = this._xmlContent.replace(current, value);
    }

    public get packageItems(): UnzippedItem[] | undefined {
        return this._packageZip && this._packageZip.zipItems;
    }

    public get rootData(): ExcelCustomXmlRootData | undefined {
        return this._rootData;
    }

    public get metaData(): ExcelCustomXmlMetadata | undefined {
        return this._metaData;
    }

    public get metaDataItems(): UnzippedItem[] | undefined {
        return this._metaDataZip && this._metaDataZip.zipItems;
    }

    public get metaDataXml(): string | undefined {
        return this._metaDataXml && this._metaDataXml[0];
    }

    public set metaDataXml(value: string) {
        if (!this._metaDataXml) {
            return;
        }
        this._metaDataXml[0] = value;
    }

    constructor(xmlContent: string) {
        this._xmlContent = xmlContent;
    }

    public async parse(): Promise<void> {
        if (this._parseResult) {
            return;
        }
        const mashupBase64 = this.mashupBase64;
        if (!mashupBase64) {
            this._parseResult = 'DataMashupNotFound';
            return;
        }
        const mashupArray = Uint8ArrayUtils.fromBase64(mashupBase64);
        if (!mashupArray) {
            this._parseResult = 'Base64DecodeError';
            return;
        }
        const rootData = ParserRoot.parse(mashupArray as never);
        if (!rootData) {
            this._parseResult = 'ParseRootError';
            return;
        }
        this._parseResult = 'Working';
        this._rootData = rootData;
        this._packageZip = new ExcelZip(rootData.packageParts);
        const packageUnzipResult = await this._packageZip.unzip();
        if (!packageUnzipResult.ok) {
            this._parseResult = 'PackageUnzipError';
            return;
        }
        const metadataArray = Uint8ArrayUtils.from(rootData.metadata);
        if (!metadataArray) {
            this._parseResult = 'MetadataError';
            return;
        }
        this._metaData = ParserMetadata.parse(metadataArray as never);
        this._metaDataZip = new ExcelZip(this._metaData.content);
        const metadataUnzipResult = await this._metaDataZip.unzip();
        if (!metadataUnzipResult.ok) {
            this._parseResult = 'MetadataUnzipError';
            return;
        }
        this._metaDataXml = Uint8ArrayUtils.toStringEncoding(
            this._metaData.metadataXml
        );
        this._parseResult = 'Success';
    }

    public resetPermissions(): void {
        if (!this.rootData) {
            return;
        }
        this.rootData.permissions = Uint8ArrayUtils.toNumberArray(
            Uint8ArrayUtils.fromString(MashupPermissionDefaults)
        );
    }

    public async save(): Promise<Result<string>> {
        if (
            !this.rootData ||
            !this.metaData ||
            !this._packageZip ||
            !this._metaDataZip ||
            !this._metaDataXml
        ) {
            return {
                ok: false,
                error: 'Unable to save because vital data is missing.',
            };
        }
        const zipResult = await this._packageZip.zip();
        if (!zipResult.ok) {
            return zipResult;
        }
        const buffers: Uint8Array[] = [];
        const { version, permissions, permissionBindings } = this.rootData;
        Uint8ArrayUtils.appendInt32LE(buffers, version);
        Uint8ArrayUtils.appendLenLE(buffers, zipResult.data);
        Uint8ArrayUtils.appendLenLE(buffers, permissions);
        Uint8ArrayUtils.appendMetadataLE(
            buffers,
            this._metaDataZip,
            this.metaData,
            Uint8ArrayUtils.fromStringEncoding(
                this._metaDataXml[0],
                this._metaDataXml[1],
                this._metaDataXml[2]
            )
        );
        Uint8ArrayUtils.appendLenLE(buffers, permissionBindings);
        const buffer = Uint8ArrayUtils.concat(buffers);
        const base64 = Uint8ArrayUtils.toBase64(buffer);
        this.mashupBase64 = base64;
        return {
            ok: true,
            data: this._xmlContent,
        };
    }
}

export class ZipFile {
    private readonly zipData: Uint8Array;
    private _zipItems: UnzippedItem[] | undefined;

    public get zipItems(): UnzippedItem[] | undefined {
        return this._zipItems;
    }

    constructor(zipData: Uint8Array) {
        this.zipData = zipData;
    }

    public async unzip(): Promise<Result<UnzippedItem[]>> {
        if (this._zipItems) {
            return {
                ok: true,
                data: this._zipItems,
            };
        }
        const data = Uint8ArrayUtils.toNumberArray(this.zipData);
        const zipResult = await Unzip(data);
        if (!zipResult.ok) {
            return zipResult;
        }
        this._zipItems = zipResult.data;
        return {
            ok: true,
            data: this._zipItems,
        };
    }

    public async zip(): Promise<Result<Uint8Array>> {
        if (!this._zipItems) {
            return {
                ok: true,
                data: this.zipData,
            };
        }
        const items = this._zipItems.map((item) => {
            if (!item.encoding) {
                return item;
            }
            const data = Uint8ArrayUtils.fromStringEncoding(
                item.data,
                item.encoding
            );
            return { ...item, data };
        });
        const zipResult = await Zip(items);
        if (!zipResult.ok) {
            return zipResult;
        }
        return {
            ok: true,
            data: zipResult.data,
        };
    }

    public setFileContents(
        item: UnzippedItem,
        data: UnzippedItem['data']
    ): Result<void> {
        const zipItem = this.zipItems && this.zipItems.find((o) => o === item);
        if (!zipItem) {
            return {
                ok: false,
                error: 'File not found in zip.',
            };
        }
        zipItem.data = data;
        return {
            ok: true,
            data: undefined,
        };
    }
}

export class ExcelZip extends ZipFile {
    private _mashupItem: UnzippedItem | null | undefined;
    private _mashupInstance: ExcelCustomXml | undefined;
    private _powerQueryItems: UnzippedItem[] | undefined;

    public get zipItems(): UnzippedItem[] {
        return super.zipItems || [];
    }

    public get mashupItem(): UnzippedItem | undefined {
        if (this._mashupItem !== undefined) {
            return this._mashupItem || undefined;
        }
        this._mashupItem = this.getMashupFile();
        return this._mashupItem;
    }

    public get mashupInstance(): ExcelCustomXml | undefined {
        if (this._mashupInstance) {
            return this._mashupInstance;
        }
        if (!this.mashupItem) {
            return undefined;
        }
        const data = this.mashupItem.data as string;
        this._mashupInstance = new ExcelCustomXml(data);
        return this._mashupInstance;
    }

    constructor(zipData: Uint8ArrayUtilsFrom) {
        zipData = Uint8ArrayUtils.from(zipData);
        super(zipData);
    }

    private convertItemToString(item: UnzippedItem): UnzippedItem {
        if (item.type !== 'File') {
            return item;
        }
        const { path } = item;
        const isXml = path.endsWith('.xml') || path.endsWith('.m');
        if (!isXml) {
            return item;
        }
        if (typeof item.data === 'string') {
            return item;
        }
        const [data, encoding, bom] = Uint8ArrayUtils.toStringEncoding(
            item.data
        );
        return { ...item, data, encoding, bom };
    }

    public async unzip(): Promise<Result<UnzippedItem[]>> {
        const result = await super.unzip();
        if (!result.ok) {
            return result;
        }
        const { data } = result;
        for (let i = 0; i < data.length; i++) {
            let item = data[i];
            item = this.convertItemToString(item);
            data[i] = item;
        }
        const mashupInstance = this.mashupInstance;
        if (mashupInstance) {
            await mashupInstance.parse();
        }
        return {
            ok: true,
            data,
        };
    }

    public async zip(): Promise<Result<Uint8Array>> {
        const mashupItem = this.mashupItem;
        const mashupInstance = this.mashupInstance;
        if (!mashupItem || !mashupInstance) {
            return super.zip();
        }
        await mashupInstance.parse();
        mashupInstance.resetPermissions();
        const saveResult = await mashupInstance.save();
        if (!saveResult.ok) {
            return saveResult;
        }
        const fileResult = this.setFileContents(mashupItem, saveResult.data);
        if (!fileResult.ok) {
            return fileResult;
        }
        return super.zip();
    }

    private getMashupFile(): UnzippedItem | undefined {
        return this.zipItems.find(({ type, path, data }) => {
            if (
                type !== 'File' ||
                typeof data !== 'string' ||
                !path.includes('customXml') ||
                !path.includes('item')
            ) {
                return false;
            }
            return data.includes(MashupBinaryPrefix);
        });
    }

    public async getPowerQueryFiles(): Promise<UnzippedItem[] | undefined> {
        if (this._powerQueryItems) {
            return this._powerQueryItems;
        }
        if (!this.mashupInstance) {
            return;
        }
        await this.mashupInstance.parse();
        const items = this.mashupInstance.metaDataItems;
        if (!items) {
            return;
        }
        this._powerQueryItems = items.filter((o) => o.path.endsWith('.m'));
        return this._powerQueryItems;
    }
}
