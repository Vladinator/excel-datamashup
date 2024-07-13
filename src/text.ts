import { Buffer } from './buffer';

let TextDecoderImpl = TextDecoder;
let TextEncoderImpl = TextEncoder;

/**
 * If the `TextDecoder` and `TextEncoder` is already available, it means we're running in the browser,
 * so we simply keep it. Otherwise, we load the Node util module.
 */
if (
    typeof TextDecoderImpl === 'undefined' ||
    typeof TextEncoderImpl === 'undefined'
) {
    const { TextDecoder, TextEncoder } = require('node:util');
    TextDecoderImpl = TextDecoder;
    TextEncoderImpl = TextEncoder;
}

/**
 * Shared `UTF-8` encoder instance to help convert a `string` to `Uint8Array`.
 */
const Encoder = new TextEncoderImpl();

/**
 * Shared `UTF-8` decoder instance to help convert a `ArrayBuffer`/`Uint8Array` to `string`.
 */
const Decoder = new TextDecoderImpl('utf-8');

/**
 * Shared `UTF-16LE` decoder instance to help convert a `ArrayBuffer`/`Uint8Array` to `string`.
 */
const DecoderUTF16LE = new TextDecoderImpl('utf-16le');

/**
 * Converts string data into a `UTF-16LE` Buffer.
 *
 * @param data `string | Uint8Array | Buffer` which isn't `UTF-16LE`
 * @returns A `UTF-16LE` encoded `Buffer`.
 */
const EncodeUTF16LE = (data: string | Uint8Array | Buffer): Buffer => {
    const str = typeof data === 'string' ? data : Decoder.decode(data);
    const length = str.length;
    const buffer = Buffer.alloc(length * 2);
    const view = new Uint16Array(
        buffer.buffer,
        buffer.byteOffset,
        buffer.length / 2
    );
    for (let i = 0; i < length; i++) {
        view[i] = str.charCodeAt(i);
    }
    return buffer;
};

export {
    TextDecoderImpl as TextDecoder,
    TextEncoderImpl as TextEncoder,
    Encoder,
    Decoder,
    DecoderUTF16LE,
    EncodeUTF16LE,
};
