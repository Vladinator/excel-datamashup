import type { UnzippedItem, UnzippedExcel } from '../types/index.d';

import { Buffer } from './buffer';
import { ParseXml } from './datamashup';
import { ReadableStream } from './stream';
import { DecoderUTF16LE, EncodeUTF16LE } from './text';
import { Zippable, unzip, zip } from 'fflate';

/**
 * Converts chunks into a buffer, then runs unzip on the whole file.
 *
 * @param chunks Partial chunks of data.
 * @param resolve Callback when the unzip finishes processing.
 */
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

/**
 * Unzip a ZIP by passing its binary data, and having its internal files and folders returned in an array.
 *
 * @param data The ZIP data.
 * @returns Array of `UnzippedItem`.
 */
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

/**
 * Zip an array of `UnzippedItem` into a ZIP binary buffer.
 *
 * @param items The files and folders as `UnzippedItem` array that you wish to ZIP up.
 * @returns The ZIP data as `Buffer`.
 */
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

/**
 * The string that signifies the DataMashup XML tag.
 *
 * This is a `UTF-8` encoded string.
 */
const DataMashupText = '<DataMashup ';

/**
 * The string that signifies the DataMashup XML tag.
 *
 * This is a `UTF-16LE` encoded string `Buffer`.
 */
const DataMashupTextAsBuffer = EncodeUTF16LE(DataMashupText);

/**
 * Find the DataMashup file from an array of files.
 *
 * @param files The files in the ZIP.
 * @returns `UnzippedItem` if the DataMashup file was found, otherwise `undefined`.
 */
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

/**
 * Convert Excel data into `UnzippedExcel` object.
 *
 * @param data The ZIP data.
 * @returns `UnzippedExcel` containing data and helper methods.
 */
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
    /**
     * Get the Power Query formula.
     *
     * @returns `string` if found, otherwise `undefined`.
     */
    function getFormula(): string | undefined {
        const { datamashup } = results;
        if (!datamashup || datamashup.error) {
            return;
        }
        return datamashup.result.getFormula();
    }
    /**
     * Set the Power Query formula.
     *
     * This also calls `resetPermissions` for you as required.
     *
     * @param formula The Power Query formula to be stored in the `Section1.m` file.
     * @returns
     */
    function setFormula(formula: string): void {
        const { datamashup } = results;
        if (!datamashup || datamashup.error) {
            return;
        }
        const { result } = datamashup;
        result.setFormula(formula);
        result.resetPermissions();
    }
    /**
     * Save the session back into an Excel file.
     *
     * @returns `Buffer` of the Excel file.
     */
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
