export * from '../types/index.d';

export * from './buffer';
export * from './stream';
export * from './text';
export * from './zip';

import { ParseXml } from './datamashup';

/* uncomment block then build for demo-purposes

import { UnzippedExcel } from '../types/index.d';

import {
    getSampleXml,
    getSampleExcelZip,
    saveSampleExcelZip,
    saveSampleXml,
} from './sample';

const RunBrowserDemo = async (excelZip: UnzippedExcel): Promise<void> => {
    const origFormula = excelZip.getFormula();
    if (!origFormula) {
        console.error('Unable to find formula.');
        return;
    }
    const newFormula = origFormula.replace(
        '"This is an example."',
        '"This is the browser demonstration."'
    );
    excelZip.setFormula(newFormula);
    await saveSampleExcelZip(excelZip);
};

const RunTerminalDemo = async (sampleXml: string): Promise<void> => {
    const result = await ParseXml(sampleXml);
    if (typeof result === 'string') {
        console.error(result);
        return;
    }
    const origFormula = result.getFormula();
    if (!origFormula) {
        console.error('Unable to find formula.');
        return;
    }
    const newFormula = origFormula.replace(
        '"This is an example."',
        '"This is the terminal demonstration."'
    );
    result.setFormula(newFormula);
    result.resetPermissions();
    const binaryString = await result.save();
    const newSampleXml = sampleXml.replace(
        /\"\>\s*(.*)\s*\<\/DataMashup\>\s*$/,
        `">${binaryString}</DataMashup>`
    );
    await saveSampleXml(newSampleXml);
};

const RunDemo = async (): Promise<void> => {
    const sampleXml = await getSampleXml();
    const excelZip = getSampleExcelZip();
    if (excelZip) {
        RunBrowserDemo(excelZip);
    } else {
        RunTerminalDemo(sampleXml);
    }
};

RunDemo();

// */

export { ParseXml };
