import type { UnzippedItem, UnzippedExcel } from '../types/index.d';

import { Buffer } from './buffer';
import { ParseXml } from './datamashup';
import { ReadableStream } from './stream';
import { DecoderUTF16LE, EncodeUTF16LE } from './text';
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

export const Unzip = (
    data: number[] | Uint8Array | Buffer
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

export const Zip = (items: UnzippedItem[]): Promise<Buffer> => {
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

const DataMashupText = '<DataMashup ';
const DataMashupTextAsBuffer = EncodeUTF16LE(DataMashupText);

const findDataMashupFile = (
    files: UnzippedItem[]
): UnzippedItem | undefined => {
    return files.find((file) => {
        if (
            file.type !== 'File' ||
            !file.path.includes('customXml') ||
            !file.path.includes('item')
        ) {
            return false;
        }
        const { data } = file;
        if (typeof data === 'string') {
            return data.includes(DataMashupText);
        }
        return data.includes(DataMashupTextAsBuffer);
    });
};

export const ExcelZip = async (
    data: number[] | Uint8Array | Buffer
): Promise<UnzippedExcel> => {
    const files = await Unzip(data);
    const file = findDataMashupFile(files);
    const results: UnzippedExcel = {
        files,
        getFormula,
        setFormula,
        save,
    };
    if (!file) {
        return results;
    }
    const { data: raw } = file;
    const xml = typeof raw === 'string' ? raw : DecoderUTF16LE.decode(raw);
    const result = await ParseXml(xml);
    if (typeof result === 'string') {
        results.datamashup = {
            file,
            xml,
            error: result,
        };
    } else {
        results.datamashup = {
            file,
            xml,
            result,
        };
    }
    return results;
    function getFormula(): string | undefined {
        const { datamashup } = results;
        if (!datamashup || datamashup.error) {
            return;
        }
        return datamashup.result.getFormula();
    }
    function setFormula(formula: string): void {
        const { datamashup } = results;
        if (!datamashup || datamashup.error) {
            return;
        }
        const { result } = datamashup;
        result.setFormula(formula);
        result.resetPermissions();
    }
    async function save(): Promise<Buffer> {
        const { datamashup } = results;
        if (!datamashup || datamashup.error) {
            return Zip(results.files);
        }
        const { result, file, xml } = datamashup;
        const binaryString = await result.save();
        const newXml = xml.replace(
            /\"\>\s*(.*)\s*\<\/DataMashup\>\s*$/,
            `">${binaryString}</DataMashup>`
        );
        datamashup.xml = newXml;
        const xmlEncoded = EncodeUTF16LE(newXml);
        file.data = xmlEncoded;
        return Zip(results.files);
    }
};
