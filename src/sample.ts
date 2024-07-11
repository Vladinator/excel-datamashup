import type { UnzippedItem } from '../types';
import { Buffer } from './buffer';
import { TextDecoder, EncodeUTF16LE } from './text';
import { Unzip, Zip } from './zip';
import demoXml from '../sample/demo.json';

const demoXmlPromise: Promise<string> = Promise.resolve(demoXml);

const userFile: {
    file?: File;
    zipFiles?: UnzippedItem[];
    zipFile?: UnzippedItem;
} = {};

const requestFile = (): Promise<string> | undefined => {
    if (typeof document === 'undefined') return;
    return new Promise((resolve) => {
        const input = document.createElement('input');
        input.type = 'file';
        input.accept = '.xlsx,.xlsm,.xlsb';
        input.multiple = false;
        input.webkitdirectory = false;
        input.style.display = 'none';
        input.addEventListener('cancel', () => output());
        input.addEventListener('change', change);
        document.body.appendChild(input);
        input.click();
        document.body.removeChild(input);
        function output(data?: string): void {
            if (data && data.length) {
                resolve(data);
            } else {
                resolve(demoXmlPromise);
            }
        }
        function change(event: any): void {
            const files: File[] | undefined = event?.target?.files;
            if (!files || !files.length) return output();
            userFile.file = files[0];
            read();
        }
        function read(): void {
            if (!userFile.file) return;
            const reader = new FileReader();
            reader.addEventListener('load', (event) => {
                const result = event.target?.result;
                if (typeof result === 'string') {
                    output(result);
                } else if (result) {
                    extract(result);
                } else {
                    output();
                }
            });
            reader.addEventListener('error', (event) => {
                const error = event.target?.error;
                if (error) {
                    console.error(error);
                }
                output();
            });
            reader.readAsArrayBuffer(userFile.file);
        }
        function extract(arrayBuffer: ArrayBuffer): void {
            const decoder = new TextDecoder('utf-16le');
            const buffer = new Uint8Array(arrayBuffer);
            Unzip(buffer)
                .then((items) => {
                    const item = items.find(
                        (item) => item.path === 'customXml/item1.xml'
                    );
                    if (item) {
                        const data =
                            typeof item.data === 'string'
                                ? item.data
                                : decoder.decode(item.data);
                        userFile.zipFiles = items;
                        userFile.zipFile = item;
                        output(data);
                    } else {
                        output();
                    }
                })
                .catch((err) => {
                    console.error(err);
                    output();
                });
        }
    });
};

const downloadFile = (buffer: Buffer, type: string, name: string): void => {
    const blob = new Blob([buffer], { type });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = name;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
};

const saveFile = (xml: string): Promise<boolean> | undefined => {
    if (typeof document === 'undefined') return;
    const { file, zipFiles, zipFile } = userFile;
    if (!file || !zipFiles || !zipFile) return;
    zipFile.data = EncodeUTF16LE(xml);
    return new Promise((resolve) => {
        Zip(zipFiles)
            .then((buffer) => {
                try {
                    downloadFile(buffer, file.type, `Modified ${file.name}`);
                    resolve(true);
                } catch (ex) {
                    console.error(ex);
                    resolve(false);
                }
            })
            .catch((err) => {
                console.error(err);
                resolve(false);
            });
    });
};

export const getSampleXml = (): Promise<string> => {
    const promise: Promise<string> | undefined = requestFile();
    if (promise) return promise;
    return demoXmlPromise;
};

export const saveSampleXml = async (xml: string): Promise<void> => {
    const success = await saveFile(xml);
    if (success) return;
    console.log(xml);
};
