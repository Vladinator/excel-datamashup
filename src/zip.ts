import { type Zippable, unzip, zip } from 'fflate';
import type { Result } from './types';
import {
    type Uint8ArrayUtilsFrom,
    type Uint8ArrayUtilsStringEncoding,
    Uint8ArrayUtils,
} from './utils';

export type UnzippedItem<T = Uint8Array> = {
    path: string;
    type: 'File' | 'Directory';
    size: number;
    data: T | string;
    encoding?: Uint8ArrayUtilsStringEncoding;
    bom?: boolean;
};

export const Unzip = (
    data: Uint8ArrayUtilsFrom
): Promise<Result<UnzippedItem[]>> => {
    const array = Uint8ArrayUtils.from(data);
    return new Promise((resolve) => {
        try {
            unzip(array, (error, files) => {
                if (error) {
                    resolve({ ok: false, error });
                    return;
                }
                const data: UnzippedItem[] = [];
                for (const [path, fileData] of Object.entries(files)) {
                    data.push({
                        path,
                        type: 'File',
                        size: fileData.length,
                        data: fileData,
                    });
                }
                resolve({ ok: true, data });
            });
        } catch (ex) {
            resolve({ ok: false, error: ex as Error });
        }
    });
};

export const Zip = (items: UnzippedItem[]): Promise<Result<Uint8Array>> => {
    return new Promise((resolve) => {
        const zippable = items.reduce((pv, cv) => {
            const data =
                typeof cv.data === 'string'
                    ? Uint8ArrayUtils.fromString(cv.data)
                    : cv.data;
            pv[cv.path] = data;
            return pv;
        }, {} as Zippable);
        try {
            zip(zippable, (error, data) => {
                if (error) {
                    resolve({ ok: false, error });
                    return;
                }
                resolve({ ok: true, data });
            });
        } catch (ex) {
            resolve({ ok: false, error: ex as Error });
        }
    });
};
