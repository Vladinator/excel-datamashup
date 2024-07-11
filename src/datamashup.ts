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
import { unzipData, zipItems } from './zip';

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

const ParserMetadata = new Parser()
    .endianness('little')
    .uint32le('version')
    .uint32le('metadataXmlLength')
    .array('metadataXml', { type: 'uint8', length: 'metadataXmlLength' })
    .uint32le('contentLength')
    .array('content', { type: 'uint8', length: 'contentLength' });

const DataMashupRegexp = /<DataMashup[^>]*>(.*?)<\/DataMashup>/s;

const FormulaSectionDefault = 'Section1.m';

const PermissionDefaults = `<?xml version="1.0" encoding="utf-8"?>\r\n<PermissionList xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">\r\n\t<CanEvaluateFuturePackages>false</CanEvaluateFuturePackages>\r\n\t<FirewallEnabled>true</FirewallEnabled>\r\n\t<WorkbookGroupType xsi:nil="true" />\r\n</PermissionList>`;

const ExtractDataMashupBinary = (xml: string): string | undefined => {
    const match = xml.match(DataMashupRegexp);
    return match ? match[1].trim() : undefined;
};

const ConvertToBuffer = (
    data: number[] | Buffer | Uint8Array
): globalThis.Buffer => {
    if (!(data instanceof Buffer) && !(data instanceof Uint8Array)) {
        data = Uint8Array.from(data);
    }
    return Buffer.from(data) as never;
};

const ConvertBinaryToBuffer = (binaryString: string): globalThis.Buffer => {
    return ConvertToBuffer(base64.toByteArray(binaryString));
};

const ConvertUnzippedItemXmlToString = (item: UnzippedItem): UnzippedItem => {
    const { path } = item;
    const isXml = path.endsWith('.xml') || path.endsWith('.m');
    if (isXml && typeof item.data !== 'string') {
        return { ...item, data: item.data.toString('utf-8') };
    }
    return item;
};

export const ParseXml = async (
    xmlContent: string
): Promise<ParseResult | ParseError> => {
    const binaryString = ExtractDataMashupBinary(xmlContent);
    if (!binaryString) return 'DataMashupNotFound';

    const binaryBuffer = ConvertBinaryToBuffer(binaryString);
    if (!binaryBuffer) return 'Base64DecodeError';

    const rootData = ParserRoot.parse(binaryBuffer);
    if (!rootData) return 'ParseRootError';

    const packageItems = await unzipData(rootData.packageParts);

    const packageParts = packageItems.map(ConvertUnzippedItemXmlToString);

    const permissions = Decoder.decode(ConvertToBuffer(rootData.permissions));

    const metadataData = ParserMetadata.parse(
        ConvertToBuffer(rootData.metadata)
    );

    const metadataXml = Decoder.decode(
        ConvertToBuffer(metadataData.metadataXml)
    );

    const metadataContentItems = await unzipData(metadataData.content);

    const metadata: Metadata = {
        version: metadataData.version,
        metadata: metadataXml,
        content: metadataContentItems,
    };

    /* TODO

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

    const getFormulaFile = (): UnzippedItem => {
        let item = result.packageParts.find((item) =>
            item.path.includes(FormulaSectionDefault)
        );
        if (!item) {
            item = {
                path: FormulaSectionDefault,
                size: 0,
                type: 'File',
                data: '',
            };
            result.packageParts.push(item);
        }
        return item;
    };

    const appendBufferInt32LE = (buffers: Buffer[], value: number): void => {
        const buffer = Buffer.alloc(4);
        buffer.writeUInt32LE(value, 0);
        buffers.push(buffer);
    };

    const appendBuffer = (
        buffers: Buffer[],
        buffer: number[] | Buffer
    ): void => {
        if (!(buffer instanceof Buffer)) {
            buffer = Buffer.from(buffer);
        }
        buffers.push(buffer);
    };

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

    function getFormula(): string {
        const item = getFormulaFile();
        return item.data as string;
    }

    function setFormula(formula: string): void {
        const item = getFormulaFile();
        item.size = formula.length;
        item.data = formula;
    }

    function resetPermissions(): void {
        result.permissions = PermissionDefaults;
    }

    async function appendBufferLenLE(
        parentBuffers: Buffer[],
        buffer: Buffer
    ): Promise<void> {
        appendBufferInt32LE(parentBuffers, buffer.length);
        appendBufferLE(parentBuffers, buffer);
    }

    async function appendMetadataBufferLE(
        parentBuffers: Buffer[]
    ): Promise<void> {
        const buffers: Buffer[] = [];
        const { version, metadata, content } = result.metadata;
        appendBufferInt32LE(buffers, version);
        await appendBufferLenLE(buffers, Buffer.from(metadata));
        await appendBufferLenLE(buffers, await zipItems(content));
        const buffer = Buffer.concat(buffers);
        appendBufferInt32LE(parentBuffers, buffer.length);
        appendBuffer(parentBuffers, buffer);
    }

    async function save(): Promise<string> {
        const buffers: Buffer[] = [];
        const { version, packageParts, permissions, permissionBindings } =
            result;
        appendBufferInt32LE(buffers, version);
        await appendBufferLenLE(buffers, await zipItems(packageParts));
        await appendBufferLenLE(buffers, Buffer.from(permissions));
        await appendMetadataBufferLE(buffers);
        await appendBufferLenLE(buffers, Buffer.from(permissionBindings)); // TODO
        const buffer = Buffer.concat(buffers);
        const binaryString = base64.fromByteArray(buffer);
        return binaryString;
    }
};
