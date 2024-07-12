import type { UnzippedExcel } from '../types';
import { Buffer } from './buffer';
import { Encoder } from './text';
import { ExcelZip } from './zip';
import demoXml from '../sample/demo.json';

const demoXmlPromise: Promise<string> = Promise.resolve(demoXml);

const userFile: {
    file?: File;
    excelZip?: UnzippedExcel;
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
        function output(result?: UnzippedExcel): void {
            userFile.excelZip = result;
            if (result && result.datamashup) {
                resolve(result.datamashup.xml);
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
                if (result) {
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
        function extract(data: string | ArrayBuffer): void {
            const buffer =
                typeof data === 'string'
                    ? Encoder.encode(data)
                    : new Uint8Array(data);
            ExcelZip(buffer)
                .then((result) => {
                    output(result);
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

const saveExcelZip = (
    excelZip: UnzippedExcel
): Promise<boolean> | undefined => {
    if (typeof document === 'undefined') return;
    const { file } = userFile;
    if (!file) return;
    return new Promise((resolve) => {
        excelZip
            .save()
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

const saveFile = (xml: string): Promise<boolean> | undefined => {
    if (typeof document === 'undefined') return;
    const { file, excelZip } = userFile;
    if (!file || !excelZip) return;
    excelZip.setFormula(xml);
    return saveExcelZip(excelZip);
};

export const getSampleXml = (): Promise<string> => {
    const promise: Promise<string> | undefined = requestFile();
    if (promise) return promise;
    return demoXmlPromise;
};

export const getSampleExcelZip = (): UnzippedExcel | undefined => {
    return userFile.excelZip;
};

export const saveSampleExcelZip = async (
    excelZip: UnzippedExcel
): Promise<void> => {
    await saveExcelZip(excelZip);
};

export const saveSampleXml = async (xml: string): Promise<void> => {
    const success = await saveFile(xml);
    if (success) return;
    console.log(xml);
};
