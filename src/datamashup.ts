import { Parser } from 'binary-parser';
import { ExcelZip } from './excel';
import type { Result } from './types';
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
const MashupPermissionDefaults = `<?xml version="1.0" encoding="utf-8"?>\r\n<PermissionList xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">\r\n\t<CanEvaluateFuturePackages>false</CanEvaluateFuturePackages>\r\n\t<FirewallEnabled>true</FirewallEnabled>\r\n\t<WorkbookGroupType xsi:nil="true" />\r\n</PermissionList>`;

/**
 * This struct matches the top-level binary stream.
 *
 * References:
 *
 * https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-qdeff/22557f6d-7c29-4554-8fe4-7b7a54ac7a2b
 */
const DataMashupRoot = new Parser()
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
const DataMashupMetadata = new Parser()
    .endianness('little')
    .uint32le('version')
    .uint32le('metadataXmlLength')
    .array('metadataXml', { type: 'uint8', length: 'metadataXmlLength' })
    .uint32le('contentLength')
    .array('content', { type: 'uint8', length: 'contentLength' });

type IDataMashupRoot = ReturnType<(typeof DataMashupRoot)['parse']>;

type IDataMashupMetadata = ReturnType<(typeof DataMashupMetadata)['parse']>;

export type IDataMashupResult = {
    root: IDataMashupRoot;
    rootPerm: Uint8Array;
    rootPermBind: Uint8Array;
    rootZip: ExcelZip;
    rootMeta: Uint8Array;
    meta: IDataMashupMetadata;
    metaZip: ExcelZip;
    metaXml: IStringEncodingInfo;
};

export class DataMashup {
    private readonly _mashup: Uint8Array;
    private readonly _data: Partial<IDataMashupResult>;
    private _parse: Promise<IDataMashupResult> | undefined;

    private constructor(mashup: Uint8Array) {
        this._mashup = mashup;
        this._data = {};
    }

    public get data(): IDataMashupResult {
        return this._data as IDataMashupResult;
    }

    private async parse(): Promise<IDataMashupResult> {
        const { data } = this;
        data.root = DataMashupRoot.parse(this._mashup as never); // Buffer extends Uint8Array, luckily the functionality added is not used by the `Parser`
        data.rootPerm = Uint8ArrayUtils.from(data.root.permissions);
        data.rootPermBind = Uint8ArrayUtils.from(data.root.permissionBindings);
        data.rootZip = await ExcelZip.unzip(data.root.packageParts);
        data.rootMeta = Uint8ArrayUtils.from(data.root.metadata);
        data.meta = DataMashupMetadata.parse(data.rootMeta as never); // Buffer extends Uint8Array, luckily the functionality added is not used by the `Parser`
        data.metaZip = await ExcelZip.unzip(data.meta.content);
        data.metaXml = Uint8ArrayUtils.toStringEncoding(data.meta.metadataXml);
        return data;
    }

    private unpack(): Promise<IDataMashupResult> {
        if (this._parse) {
            return this._parse;
        }
        this._parse = this.parse();
        return this._parse;
    }

    private async packRoot(): Promise<Uint8Array | undefined> {
        const { data } = this;
        const metadata = await this.packMeta();
        if (!metadata) {
            return;
        }
        const packageParts = await PromiseData(data.rootZip.zip());
        if (!packageParts) {
            return;
        }
        const permissions = data.rootPerm;
        const permissionBindings = data.rootPermBind;
        const version = data.root.version;
        const totalLength =
            // version
            4 +
            // packageParts
            4 +
            packageParts.length +
            // permissions
            4 +
            permissions.length +
            // metadata
            4 +
            metadata.length +
            // permissionBindings
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

    private async packMeta(): Promise<Uint8Array | undefined> {
        const { data } = this;
        const content = await PromiseData(data.metaZip.zip());
        if (!content) {
            return;
        }
        const metadataXml = Uint8ArrayUtils.fromStringEncodingInfo(
            data.metaXml
        );
        const version = data.meta.version;
        const totalLength =
            // version
            4 +
            // metadataXml
            4 +
            metadataXml.length +
            // content
            4 +
            content.length;
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
        const array = await this.packRoot();
        return array;
    }

    public get rootItems(): UnzippedItem[] {
        return this.data.rootZip.zipItems;
    }

    public get metaItems(): UnzippedItem[] {
        return this.data.metaZip.zipItems;
    }

    public setFileContents(
        item: UnzippedItem,
        data: UnzippedItem['data']
    ): Result<void> {
        if (this.rootItems.some((o) => o === item)) {
            return this.data.rootZip.setFileContents(item, data);
        }
        if (this.metaItems.some((o) => o === item)) {
            return this.data.metaZip.setFileContents(item, data);
        }
        return { ok: false, error: 'File not found.' };
    }

    public resetPermissions(): void {
        this.data.rootPerm = Uint8ArrayUtils.fromString(
            MashupPermissionDefaults
        );
    }

    public static async unpack(mashup: Uint8Array): Promise<DataMashup> {
        const instance = new this(mashup);
        await instance.unpack();
        return instance;
    }
}
