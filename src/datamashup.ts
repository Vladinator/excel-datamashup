import type {
    ParseError,
    UnzippedItem,
    Metadata,
    ParseResult,
} from '../types/index.d';

import base64 from 'base64-js';
import { Parser } from 'binary-parser';
import { Buffer } from './buffer';
import { crypto } from './crypto';
import { Decoder, Encoder } from './text';
import { Unzip, Zip } from './zip';

/**
 * This struct matches the top-level binary stream.
 *
 * References:
 *
 * https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-qdeff/22557f6d-7c29-4554-8fe4-7b7a54ac7a2b
 */
const ParserRoot = new Parser()
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
const ParserMetadata = new Parser()
    .endianness('little')
    .uint32le('version')
    .uint32le('metadataXmlLength')
    .array('metadataXml', { type: 'uint8', length: 'metadataXmlLength' })
    .uint32le('contentLength')
    .array('content', { type: 'uint8', length: 'contentLength' });

/**
 * Matching RegExp to extract the DataMashup XML tag and the base64 binary data.
 */
const DataMashupRegexp = /<DataMashup[^>]*>(.*?)<\/DataMashup>/s;

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
const FormulaSectionDefault = 'Section1.m';

/**
 * If the top-level binary stream permission bindings become cryptographically invalid, then
 * we need to reset the permissions XML to this default to indicate that content has changed outside of Excel.
 *
 * References:
 *
 * https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-qdeff/d0959ba8-ac8d-4bee-bb58-9a869d7b226a
 */
const PermissionDefaults = `<?xml version="1.0" encoding="utf-8"?>\r\n<PermissionList xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">\r\n\t<CanEvaluateFuturePackages>false</CanEvaluateFuturePackages>\r\n\t<FirewallEnabled>true</FirewallEnabled>\r\n\t<WorkbookGroupType xsi:nil="true" />\r\n</PermissionList>`;

/**
 * From a XML document find the DataMashup tag and extract the base64 binary string.
 *
 * @param xml The `customXml\item1.xml` file contents.
 * @returns base64 binary `string` if found, otherwise `undefined`
 */
const ExtractDataMashupBinary = (xml: string): string | undefined => {
    const match = xml.match(DataMashupRegexp);
    return match ? match[1].trim() : undefined;
};

/**
 * Takes binary data and converts it to a Buffer.
 *
 * @param data `number` array, `Uint8Array` (or another `Buffer`) will be converted to a proper `Buffer`
 * @returns `Buffer`
 */
const ConvertToBuffer = (
    data: number[] | Buffer | Uint8Array
): globalThis.Buffer => {
    if (!(data instanceof Buffer) && !(data instanceof Uint8Array)) {
        data = Uint8Array.from(data);
    }
    return Buffer.from(data) as never;
};

/**
 * Takes base64 binary string and converts it to a Buffer.
 *
 * @param binaryString
 * @returns `Buffer`
 */
const ConvertBinaryToBuffer = (binaryString: string): globalThis.Buffer => {
    return ConvertToBuffer(base64.toByteArray(binaryString));
};

/**
 * The provided unzipped item will have its `data` property converted from `Buffer` to a `string`.
 *
 * This only affects `.xml` and `.m` files.
 *
 * This alters the incoming `item.data` property on `item`.
 *
 * @param item `UnzippedItem`
 * @returns The same `UnzippedItem` `item` as passed to the function.
 */
const ConvertUnzippedItemXmlToString = (item: UnzippedItem): UnzippedItem => {
    const { path } = item;
    const isXml = path.endsWith('.xml') || path.endsWith('.m');
    if (isXml && typeof item.data !== 'string') {
        return { ...item, data: item.data.toString('utf-8') };
    }
    return item;
};

/**
 * Pass the `customXml\item1.xml` content and it will be processed into `ParseResult | ParseError` object.
 *
 * You need to control if the returned object is a `string` and thus a `ParseError`, or the usable `ParseResult` object.
 *
 * @param xmlContent The `customXml\item1.xml` content.
 * @returns `ParseResult | ParseError` object.
 */
export const ParseXml = async (
    xmlContent: string
): Promise<ParseResult | ParseError> => {
    const binaryString = ExtractDataMashupBinary(xmlContent);
    if (!binaryString) return 'DataMashupNotFound';

    const binaryBuffer = ConvertBinaryToBuffer(binaryString);
    if (!binaryBuffer) return 'Base64DecodeError';

    const rootData = ParserRoot.parse(binaryBuffer);
    if (!rootData) return 'ParseRootError';

    const packageItems = await Unzip(rootData.packageParts);

    const packageParts = packageItems.map(ConvertUnzippedItemXmlToString);

    const permissions = Decoder.decode(ConvertToBuffer(rootData.permissions));

    const metadataData = ParserMetadata.parse(
        ConvertToBuffer(rootData.metadata)
    );

    const metadataXml = Decoder.decode(
        ConvertToBuffer(metadataData.metadataXml)
    );

    const metadataContentItems = await Unzip(metadataData.content);

    const metadata: Metadata = {
        version: metadataData.version,
        metadata: metadataXml,
        content: metadataContentItems,
    };

    /*
    // TODO: need to use the crypto file to handle this, but
    // it might not be feasible as originally it needs DPAPI, which
    // might be OS exclusive and not available in the browser.
    // https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-qdeff/d0959ba8-ac8d-4bee-bb58-9a869d7b226a
    // https://learn.microsoft.com/en-us/previous-versions/ms995355(v%3Dmsdn.10)

    const sha256 = async (data: Uint8Array): Promise<Uint8Array> => {
        return new Uint8Array(await crypto.subtle.digest('SHA-256', data));
    };

    const concatBuffers = (...buffers: Uint8Array[]): Uint8Array => {
        const length = buffers.reduce((acc, buffer) => acc + buffer.length, 0);
        const buffer = new Uint8Array(length);
        let offset = 0;
        for (const buffer of buffers) {
            buffer.set(buffer, offset);
            offset += buffer.length;
        }
        return buffer;
    };

    const prefixHash = (hash: Uint8Array): Uint8Array => {
        const lengthBuffer = new Uint8Array(4);
        new DataView(lengthBuffer.buffer).setUint32(0, hash.length, true);
        const buffer = new Uint8Array(lengthBuffer.length + hash.length);
        buffer.set(lengthBuffer, 0);
        buffer.set(hash, lengthBuffer.length);
        return buffer;
    };

    const getScopeEntropy = (
        scope?: string,
        entropy?: Uint8Array
    ): [string, Uint8Array] => {
        scope ||= 'Current user';
        entropy ||= Encoder.encode('DataExplorer Package Components');
        return [scope, entropy];
    };

    const deriveKey = async (
        entropy: Uint8Array,
        scope: string
    ): Promise<CryptoKey> => {
        const scopeBuffer = Encoder.encode(scope);
        const keyMaterial = await crypto.subtle.importKey(
            'raw',
            entropy,
            { name: 'PBKDF2' },
            false,
            ['deriveBits', 'deriveKey']
        );
        const key = await crypto.subtle.deriveKey(
            {
                name: 'PBKDF2',
                salt: scopeBuffer,
                iterations: 100000,
                hash: 'SHA-256',
            },
            keyMaterial,
            { name: 'AES-GCM', length: 256 },
            true,
            ['encrypt', 'decrypt']
        );
        return key;
    };

    const generateIV = () => crypto.getRandomValues(new Uint8Array(12));

    const encryptData = async (
        data: Uint8Array,
        key: CryptoKey,
        iv: Uint8Array
    ): Promise<{
        encryptedData: Uint8Array;
        cipherText: Uint8Array;
        authTag: Uint8Array;
    }> => {
        const encryptedDataArray = await crypto.subtle.encrypt(
            {
                name: 'AES-GCM',
                iv,
            },
            key,
            data
        );
        const encryptedData = new Uint8Array(encryptedDataArray);
        const ciphertextLength = data.length;
        const cipherText = encryptedData.slice(0, ciphertextLength);
        const authTag = encryptedData.slice(ciphertextLength);
        return { encryptedData, cipherText, authTag };
    };

    const encrypt = async (
        packageParts: Uint8Array,
        permissions: Uint8Array,
        scope?: string,
        entropy?: Uint8Array
    ): Promise<Uint8Array> => {
        [scope, entropy] = getScopeEntropy(scope, entropy);

        const packagePartsHash = await sha256(packageParts);
        const permissionsHash = await sha256(permissions);

        const packagePartsHashPrefixed = prefixHash(packagePartsHash);
        const permissionsHashPrefixed = prefixHash(permissionsHash);

        const hashes = concatBuffers(
            packagePartsHashPrefixed,
            permissionsHashPrefixed
        );

        const key = await deriveKey(entropy, scope);
        const iv = generateIV();
        const { encryptedData, authTag } = await encryptData(hashes, key, iv);

        return concatBuffers(iv, authTag, encryptedData);
    };

    const decrypt = async (
        encryptedData: Uint8Array,
        scope?: string,
        entropy?: Uint8Array
    ): Promise<Uint8Array> => {
        [scope, entropy] = getScopeEntropy(scope, entropy);

        const iv = encryptedData.subarray(0, 12);
        const authTag = encryptedData.subarray(12, 28);
        const data = encryptedData.subarray(28);

        const keyMaterial = await crypto.subtle.importKey(
            'raw',
            entropy,
            { name: 'PBKDF2' },
            false,
            ['deriveBits', 'deriveKey']
        );

        const key = await crypto.subtle.deriveKey(
            {
                name: 'PBKDF2',
                salt: Encoder.encode(scope),
                iterations: 100000,
                hash: 'SHA-256',
            },
            keyMaterial,
            { name: 'AES-GCM', length: 256 },
            true,
            ['encrypt', 'decrypt']
        );

        const decryptedData = await crypto.subtle.decrypt(
            {
                name: 'AES-GCM',
                iv,
                tag: authTag,
            } as AesGcmParams,
            key,
            data
        );

        return new Uint8Array(decryptedData);
    };

    const permissionBindings = await decrypt(
        Uint8Array.from(rootData.permissionBindings)
    );

    const permissionBindingsTest = await encrypt(
        permissionBindings,
        Uint8Array.from(rootData.permissions)
    );

    console.log(permissionBindings); // TODO
    console.log(permissionBindingsTest); // TODO

    // */

    /**
     * Grabs the formula file from the correct package parts item.
     *
     * @returns `UnzippedItem` if available, otherwise `undefined`
     */
    const getFormulaFile = (): UnzippedItem | undefined => {
        return result.packageParts.find((item) =>
            item.path.includes(FormulaSectionDefault)
        );
    };

    /**
     * Appends the value as a Int32 little-endian to the existing buffer array.
     *
     * @param buffers The buffer array.
     * @param value The number to Int32 little-endian and store.
     */
    const appendBufferInt32LE = (buffers: Buffer[], value: number): void => {
        const buffer = Buffer.alloc(4);
        buffer.writeUInt32LE(value, 0);
        buffers.push(buffer);
    };

    /**
     * Appends the buffer to the buffer array.
     *
     * Will not manipulate the byte-order.
     *
     * @param buffers The buffer array.
     * @param buffer A number array or another Buffer.
     */
    const appendBuffer = (
        buffers: Buffer[],
        buffer: number[] | Buffer
    ): void => {
        if (!(buffer instanceof Buffer)) {
            buffer = Buffer.from(buffer);
        }
        buffers.push(buffer);
    };

    /**
     * Appends the buffer to the buffer array.
     *
     * Will manipulate the byte-order to little-endian.
     *
     * @param buffers The buffer array.
     * @param buffer A string, number array or another Buffer.
     */
    const appendBufferLE = (
        buffers: Buffer[],
        buffer: string | number[] | Buffer
    ): void => {
        if (typeof buffer === 'string') {
            buffer = Buffer.from(buffer);
        }
        const bufferLE = Buffer.alloc(buffer.length);
        for (let i = 0; i < buffer.length; i++) {
            bufferLE.writeUInt8(buffer[i], i);
        }
        buffers.push(bufferLE);
    };

    /**
     * The `ParseResult` object which contains all the data and helper methods.
     */
    const result: ParseResult = {
        version: rootData.version,
        packageParts,
        permissions,
        metadata,
        permissionBindings: rootData.permissionBindings, // TODO,
        getFormula,
        setFormula,
        resetPermissions,
        save,
    };

    return result;

    /**
     * Retrieve the Power Query formula.
     *
     * @returns `string` if available, otherwise `undefined`
     */
    function getFormula(): string | undefined {
        const item = getFormulaFile();
        if (!item) return;
        return item.data as string;
    }

    /**
     * Set the Power Query formula.
     *
     * @param formula The desired Power Query to store.
     */
    function setFormula(formula: string): void {
        const item = getFormulaFile();
        if (!item) return;
        item.size = formula.length;
        item.data = formula;
    }

    /**
     * Resets the permissions to defaults.
     *
     * This must be called if the formula is changed.
     */
    function resetPermissions(): void {
        result.permissions = PermissionDefaults;
    }

    /**
     * Appends a buffer with a length prefix entry.
     *
     * The length is Int32 little-endian encoded, right before the Buffer itself is little-endian encoded and stored into the buffer array.
     *
     * @param parentBuffers The buffer array.
     * @param buffer The buffer we wish to store.
     */
    async function appendBufferLenLE(
        parentBuffers: Buffer[],
        buffer: Buffer
    ): Promise<void> {
        appendBufferInt32LE(parentBuffers, buffer.length);
        appendBufferLE(parentBuffers, buffer);
    }

    /**
     * Appends the metadata buffer as little-endian.
     *
     * @param parentBuffers The buffer array.
     */
    async function appendMetadataBufferLE(
        parentBuffers: Buffer[]
    ): Promise<void> {
        const buffers: Buffer[] = [];
        const { version, metadata, content } = result.metadata;
        appendBufferInt32LE(buffers, version);
        await appendBufferLenLE(buffers, Buffer.from(metadata));
        await appendBufferLenLE(buffers, await Zip(content));
        const buffer = Buffer.concat(buffers);
        appendBufferInt32LE(parentBuffers, buffer.length);
        appendBuffer(parentBuffers, buffer);
    }

    /**
     * Saves the current `ParseResult` state back into base64 binary `string`.
     *
     * @returns base64 binary `string`
     */
    async function save(): Promise<string> {
        const buffers: Buffer[] = [];
        const { version, packageParts, permissions, permissionBindings } =
            result;
        appendBufferInt32LE(buffers, version);
        await appendBufferLenLE(buffers, await Zip(packageParts));
        await appendBufferLenLE(buffers, Buffer.from(permissions));
        await appendMetadataBufferLE(buffers);
        await appendBufferLenLE(buffers, Buffer.from(permissionBindings)); // TODO: DPAPI not implemented so we just re-use the original data
        const buffer = Buffer.concat(buffers);
        const binaryString = base64.fromByteArray(buffer);
        return binaryString;
    }
};
