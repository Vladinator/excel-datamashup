import base64 from 'base64-js';
import { type ExcelCustomXmlMetadata, ExcelZip } from './excel';

/** The supported data types to convert into a `Uint8Array` which the `Uint8ArrayUtils.from` supports. */
export type Uint8ArrayUtilsFrom = Uint8Array | number[];

export type Uint8ArrayUtilsTo = Uint8ArrayUtilsFrom | ArrayBuffer;

export type Uint8ArrayUtilsStringEncoding = 'utf-8' | 'utf-16le' | 'utf-16be';

const Uint8ArrayUtilsStringEncoding: Record<
    Uint8ArrayUtilsStringEncoding,
    Uint8ArrayUtilsStringEncoding
> = {
    'utf-8': 'utf-8',
    'utf-16le': 'utf-16le',
    'utf-16be': 'utf-16be',
} as const;

/** In the browser environment this can only be `utf-8`. */
const textEncoder = new TextEncoder();

const textDecoder = new TextDecoder(Uint8ArrayUtilsStringEncoding['utf-8']);

const textDecoders = [
    new TextDecoder(Uint8ArrayUtilsStringEncoding['utf-8'], { fatal: true }),
    new TextDecoder(Uint8ArrayUtilsStringEncoding['utf-16le'], { fatal: true }),
    new TextDecoder(Uint8ArrayUtilsStringEncoding['utf-16be'], { fatal: true }),
];

export class Uint8ArrayUtils {
    /** Convert any supported data type into `Uint8Array`. */
    public static from(data: Uint8ArrayUtilsFrom): Uint8Array {
        if (Array.isArray(data)) {
            data = Uint8Array.from(data);
        }
        return data;
    }

    /** Convert a base64 string into `Uint8Array`. */
    public static fromBase64(base64String: string): Uint8Array {
        return base64.toByteArray(base64String);
    }

    /** Convert a string into `Uint8Array`. */
    public static fromString(data: string | Uint8Array): Uint8Array {
        if (data instanceof Uint8Array) {
            return data;
        }
        return textEncoder.encode(data);
    }

    public static fromStringEncoding(
        data: string | Uint8Array,
        encoding: Uint8ArrayUtilsStringEncoding,
        addBOM?: boolean
    ): Uint8Array {
        switch (encoding) {
            case 'utf-8': {
                const array = this.fromString(data);
                if (!addBOM) {
                    return array;
                }
                const bom = new Uint8Array([0xef, 0xbb, 0xbf]);
                const result = new Uint8Array(bom.length + array.length);
                result.set(bom, 0);
                result.set(array, bom.length);
                return result;
            }
            case 'utf-16le':
                return this.encodeUTF16LE(data, addBOM);
            case 'utf-16be':
                return this.encodeUTF16BE(data, addBOM);
        }
    }

    public static toBase64(data: Uint8ArrayUtilsFrom): string {
        data = this.from(data);
        return base64.fromByteArray(data);
    }

    public static toString(data: Uint8ArrayUtilsTo): string {
        if (!(data instanceof ArrayBuffer)) {
            data = this.from(data);
        }
        return textDecoder.decode(data);
    }

    public static toStringEncoding(
        data: Uint8ArrayUtilsTo
    ): [string, Uint8ArrayUtilsStringEncoding, boolean] {
        if (!(data instanceof ArrayBuffer)) {
            data = this.from(data);
        }
        const array =
            data instanceof ArrayBuffer ? new Uint8Array(data, 0, 3) : data;
        let hasBOM = false;
        if (
            (array[0] === 0xef && array[1] === 0xbb && array[2] === 0xbf) || // utf-8
            (array[0] === 0xfe && array[1] === 0xff) || // utf-16le
            (array[0] === 0xff && array[1] === 0xfe) // utf-16be
        ) {
            hasBOM = true;
        }
        for (const decoder of textDecoders) {
            try {
                const decoded = decoder.decode(data);
                return [
                    decoded,
                    decoder.encoding as Uint8ArrayUtilsStringEncoding,
                    hasBOM,
                ];
            } catch {
                // console.error(ex);
            }
        }
        return [
            textDecoder.decode(data),
            textDecoder.encoding as Uint8ArrayUtilsStringEncoding,
            hasBOM,
        ];
    }

    /** Convert any supported data type into `number[]`. */
    public static toNumberArray(data: Uint8ArrayUtilsTo): number[] {
        if (Array.isArray(data)) {
            return data;
        }
        if (data instanceof ArrayBuffer) {
            data = new Uint8Array(data);
        }
        if (!(data instanceof Uint8Array)) {
            return [];
        }
        data = Array.from(data.values());
        return data;
    }

    /** Concat an array of `Uint8Array` into a flat `Uint8Array`. */
    public static concat(arrays: Uint8Array[]): Uint8Array {
        const totalLength = arrays.reduce((sum, arr) => sum + arr.length, 0);
        const result = new Uint8Array(totalLength);
        let offset = 0;
        for (const array of arrays) {
            result.set(array, offset);
            offset += array.length;
        }
        return result;
    }

    public static append(
        arrays: Uint8Array[],
        value: Uint8ArrayUtilsFrom
    ): void {
        const array = this.from(value);
        arrays.push(array);
    }

    public static appendInt32LE(arrays: Uint8Array[], value: number): void {
        const arrayBuffer = new ArrayBuffer(4);
        const view = new DataView(arrayBuffer);
        view.setUint32(0, value, true);
        const array = new Uint8Array(arrayBuffer);
        arrays.push(array);
    }

    public static appendLE(
        arrays: Uint8Array[],
        value: string | number[] | Uint8Array
    ): void {
        if (typeof value === 'string') {
            value = this.fromString(value);
        }
        const array = new Uint8Array(value.length);
        array.set(value);
        arrays.push(array);
    }

    public static appendLenLE(
        arrays: Uint8Array[],
        value: Uint8ArrayUtilsFrom
    ): void {
        const array = this.from(value);
        this.appendInt32LE(arrays, array.length);
        this.appendLE(arrays, array);
    }

    public static async appendMetadataLE(
        arrays: Uint8Array[],
        zip: ExcelZip,
        metadata: ExcelCustomXmlMetadata,
        metadataXmlArray: Uint8Array
    ): Promise<void> {
        const zipResult = await zip.zip();
        if (!zipResult.ok) {
            return;
        }
        const buffers: Uint8Array[] = [];
        const { version } = metadata;
        this.appendInt32LE(buffers, version);
        this.appendLenLE(buffers, metadataXmlArray);
        this.appendLenLE(buffers, zipResult.data);
        const buffer = this.concat(buffers);
        this.appendInt32LE(arrays, buffer.length);
        this.append(arrays, buffer);
    }

    private static encodeUTF16(
        data: string | Uint8ArrayUtilsFrom,
        isLe: boolean,
        addBOM?: boolean
    ): Uint8Array {
        const str = typeof data === 'string' ? data : this.toString(data);
        const length = str.length;
        const buffer = new Uint8Array(length * 2 + (addBOM ? 2 : 0));
        let offset = 0;
        if (addBOM) {
            if (isLe) {
                buffer[0] = 0xff;
                buffer[1] = 0xfe;
            } else {
                buffer[0] = 0xfe;
                buffer[1] = 0xff;
            }
            offset = 2;
        }
        for (let i = 0; i < length; i++) {
            const code = str.charCodeAt(i);
            const pos = offset + i * 2;
            if (isLe) {
                buffer[pos] = code & 0xff;
                buffer[pos + 1] = code >> 8;
            } else {
                buffer[pos] = code >> 8;
                buffer[pos + 1] = code & 0xff;
            }
        }
        return buffer;
    }

    public static encodeUTF16LE(
        data: string | Uint8ArrayUtilsFrom,
        addBOM?: boolean
    ): Uint8Array {
        return this.encodeUTF16(data, true, addBOM);
    }

    public static encodeUTF16BE(
        data: string | Uint8ArrayUtilsFrom,
        addBOM?: boolean
    ): Uint8Array {
        return this.encodeUTF16(data, false, addBOM);
    }
}
