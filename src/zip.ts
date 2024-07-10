import type { UnzippedItem } from '../types';
import { Buffer } from './buffer';
import { ReadableStream } from './stream';
import { Zippable, unzip, zip } from 'fflate';

const unzipChunks = (
    chunks: number[],
    resolve: (items: UnzippedItem[]) => void
): void => {
    const buffer = new Uint8Array(chunks);
    unzip(buffer, (err, files) => {
        const items: UnzippedItem[] = [];
        if (err) {
            console.error(err);
            return resolve(items);
        }
        for (const [path, fileData] of Object.entries(files)) {
            items.push({
                path,
                type: 'File',
                size: fileData.length,
                data: Buffer.from(fileData),
            });
        }
        resolve(items);
    });
};

export const unzipData = (
    data: number[] | Uint8Array
): Promise<UnzippedItem[]> => {
    const buffer = data instanceof Uint8Array ? data : Uint8Array.from(data);
    const stream = ReadableStream.from(buffer);
    const reader = stream.getReader();
    return new Promise((resolve) => {
        const chunks: number[] = [];
        reader.read().then(function processChunk({ done, value }) {
            if (done) {
                return unzipChunks(chunks, resolve);
            }
            chunks.push(value);
            reader.read().then(processChunk);
        });
    });
};

export const zipItems = (items: UnzippedItem[]): Promise<Buffer> => {
    return new Promise((resolve) => {
        const zippable: Zippable = {};
        for (const item of items) {
            const data =
                typeof item.data === 'string'
                    ? Buffer.from(item.data)
                    : item.data;
            zippable[item.path] = data;
        }
        zip(zippable, (err, data) => {
            if (err) {
                console.error(err);
                resolve(Buffer.alloc(0));
                return;
            }
            const buffer = Buffer.from(data);
            resolve(buffer);
        });
    });
};
