import { Parser } from 'binary-parser';
import { ExcelZip } from './excel';
import { type UnzippedItem } from './zip';
import {
    type IStringEncodingInfo,
    PromiseData,
    Uint8ArrayUtils,
} from './utils';

/** Matching RegExp to extract the DataMashup XML tag and the base64 binary data. */
export const MashupBinaryRegExp = /<DataMashup[^>]*>(.*?)<\/DataMashup>/s;

/** The string that signifies the DataMashup XML tag. */
export const MashupBinaryPrefix = '<DataMashup ';

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
export const MashupFormulaSectionDefault = 'Section1.m';

/**
 * If the top-level binary stream permission bindings become cryptographically invalid, then
 * we need to reset the permissions XML to this default to indicate that content has changed outside of Excel.
 *
 * References:
 *
 * https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-qdeff/d0959ba8-ac8d-4bee-bb58-9a869d7b226a
 */
export const MashupPermissionDefaults = `<?xml version="1.0" encoding="utf-8"?>\r\n<PermissionList xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">\r\n\t<CanEvaluateFuturePackages>false</CanEvaluateFuturePackages>\r\n\t<FirewallEnabled>true</FirewallEnabled>\r\n\t<WorkbookGroupType xsi:nil="true" />\r\n</PermissionList>`;

/**
 * This struct matches the top-level binary stream.
 *
 * References:
 *
 * https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-qdeff/22557f6d-7c29-4554-8fe4-7b7a54ac7a2b
 */
export const DataMashupRoot = new Parser()
    .endianness('little')
    .uint32le('version')
    .uint32le('packagePartsLength')
    .array('packageParts', { type: 'uint8', length: 'packagePartsLength' })
    .uint32le('permissionsLength')
    .array('permissions', { type: 'uint8', length: 'permissionsLength' })
    .uint32le('metadataLength')
    .array('metadata', { type: 'uint8', length: 'metadataLength' })
    .uint32le('permissionBindingsLength')
    .array('permissionBindings', {
        type: 'uint8',
        length: 'permissionBindingsLength',
    });

/**
 * This struct matches the metadata stream contained within the top-level binary stream.
 *
 * References:
 *
 * https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-qdeff/778afc2c-02b2-4d91-aa30-52a6067b8cb9
 */
export const DataMashupMetadata = new Parser()
    .endianness('little')
    .uint32le('version')
    .uint32le('metadataXmlLength')
    .array('metadataXml', { type: 'uint8', length: 'metadataXmlLength' })
    .uint32le('contentLength')
    .array('content', { type: 'uint8', length: 'contentLength' });

export type IDataMashupRoot = ReturnType<(typeof DataMashupRoot)['parse']>;

export type IDataMashupMetadata = ReturnType<
    (typeof DataMashupMetadata)['parse']
>;

export type IDataMashupResult = {
    root: NonNullable<DataMashup['_root']>;
    rootPerm: NonNullable<DataMashup['_rootPerm']>;
    rootPermBind: NonNullable<DataMashup['_rootPermBind']>;
    rootZip: NonNullable<DataMashup['_rootZip']>;
    rootZipItems: NonNullable<DataMashup['_rootZipItems']>;
    rootMeta: NonNullable<DataMashup['_rootMeta']>;
    meta: NonNullable<DataMashup['_meta']>;
    metaZip: NonNullable<DataMashup['_metaZip']>;
    metaZipItems: NonNullable<DataMashup['_metaZipItems']>;
    metaXml: NonNullable<DataMashup['_metaXml']>;
};

export class DataMashup {
    private readonly _mashup: Uint8Array;
    private _root: IDataMashupRoot | undefined;
    private _rootPerm: Uint8Array | undefined;
    private _rootPermBind: Uint8Array | undefined;
    private _rootZip: ExcelZip | undefined;
    private _rootZipItems: UnzippedItem[] | undefined;
    private _rootMeta: Uint8Array | undefined;
    private _meta: IDataMashupMetadata | undefined;
    private _metaZip: ExcelZip | undefined;
    private _metaZipItems: UnzippedItem[] | undefined;
    private _metaXml: IStringEncodingInfo | undefined;
    private _parse: Promise<IDataMashupResult> | undefined;

    constructor(mashup: Uint8Array) {
        this._mashup = mashup;
    }

    private getResult(): IDataMashupResult {
        const result: Partial<IDataMashupResult> = {
            root: this._root,
            rootPerm: this._rootPerm,
            rootPermBind: this._rootPermBind,
            rootZip: this._rootZip,
            rootZipItems: this._rootZipItems || [],
            rootMeta: this._rootMeta,
            meta: this._meta,
            metaZip: this._metaZip,
            metaZipItems: this._metaZipItems || [],
            metaXml: this._metaXml,
        };
        return result as IDataMashupResult;
    }

    private async parse(): Promise<IDataMashupResult> {
        this._root = DataMashupRoot.parse(this._mashup as never);
        this._rootPerm = Uint8ArrayUtils.from(this._root.permissions);
        this._rootPermBind = Uint8ArrayUtils.from(
            this._root.permissionBindings
        );
        this._rootZip = new ExcelZip(this._root.packageParts);
        this._rootZipItems = await PromiseData(this._rootZip.unzip());
        this._rootMeta = Uint8ArrayUtils.from(this._root.metadata);
        this._meta = DataMashupMetadata.parse(this._rootMeta as never);
        this._metaZip = new ExcelZip(this._meta.content);
        this._metaZipItems = await PromiseData(this._metaZip.unzip());
        this._metaXml = Uint8ArrayUtils.toStringEncoding(
            this._meta.metadataXml
        );
        return this.getResult();
    }

    public unpack(): Promise<IDataMashupResult> {
        if (this._parse) {
            return this._parse;
        }
        this._parse = this.parse();
        return this._parse;
    }

    private async packRoot(
        result: IDataMashupResult
    ): Promise<Uint8Array | undefined> {
        const metadata = await this.packMeta(result);
        if (!metadata) {
            return;
        }
        const packageParts = await PromiseData(result.rootZip.zip());
        if (!packageParts) {
            return;
        }
        const permissions = result.rootPerm;
        const permissionBindings = result.rootPermBind;
        const version = result.root.version;
        if (!packageParts || !metadata) {
            return;
        }
        const totalLength =
            4 +
            4 +
            packageParts.length +
            4 +
            permissions.length +
            4 +
            metadata.length +
            4 +
            permissionBindings.length;
        const buffer = new ArrayBuffer(totalLength);
        const view = new DataView(buffer);
        const array = new Uint8Array(buffer);
        let offset = 0;
        view.setUint32(offset, version, true);
        offset += 4;
        view.setUint32(offset, packageParts.length, true);
        offset += 4;
        array.set(packageParts, offset);
        offset += packageParts.length;
        view.setUint32(offset, permissions.length, true);
        offset += 4;
        array.set(permissions, offset);
        offset += permissions.length;
        view.setUint32(offset, metadata.length, true);
        offset += 4;
        array.set(metadata, offset);
        offset += metadata.length;
        view.setUint32(offset, permissionBindings.length, true);
        offset += 4;
        array.set(permissionBindings, offset);
        // offset += permissionBindings.length;
        return array;
    }

    private async packMeta(
        result: IDataMashupResult
    ): Promise<Uint8Array | undefined> {
        const content = await PromiseData(result.metaZip.zip());
        if (!content) {
            return;
        }
        const metadataXml = Uint8ArrayUtils.fromStringEncodingInfo(
            result.metaXml
        );
        const version = result.meta.version;
        const totalLength = 4 + 4 + metadataXml.length + 4 + content.length;
        const buffer = new ArrayBuffer(totalLength);
        const view = new DataView(buffer);
        const array = new Uint8Array(buffer);
        let offset = 0;
        view.setUint32(offset, version, true);
        offset += 4;
        view.setUint32(offset, metadataXml.length, true);
        offset += 4;
        array.set(metadataXml, offset);
        offset += metadataXml.length;
        view.setUint32(offset, content.length, true);
        offset += 4;
        array.set(content, offset);
        // offset += content.length;
        return array;
    }

    public async pack(): Promise<Uint8Array | undefined> {
        await this.unpack();
        const result = this.getResult();
        const root = await this.packRoot(result);
        return root;
    }

    public resetPermissions(): void {
        this._rootPerm = Uint8ArrayUtils.fromString(MashupPermissionDefaults);
    }

    public get metaItems(): UnzippedItem[] | undefined {
        return this._metaZipItems;
    }
}
