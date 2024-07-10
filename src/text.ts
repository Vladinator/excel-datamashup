import { Buffer } from './buffer';

let TextDecoderImpl = TextDecoder;
let TextEncoderImpl = TextEncoder;

if (
    typeof TextDecoderImpl === 'undefined' ||
    typeof TextEncoderImpl === 'undefined'
) {
    const { TextDecoder, TextEncoder } = require('node:util');
    TextDecoderImpl = TextDecoder;
    TextEncoderImpl = TextEncoder;
}

const Encoder = new TextEncoderImpl();

const Decoder = new TextDecoderImpl('utf-8');

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
    EncodeUTF16LE,
};
