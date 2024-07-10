export * from '../types';

import { ParseXml } from './datamashup';

/* uncomment block then build for demo-purposes

import { getSampleXml, saveSampleXml } from './sample';

const RunDemo = async (): Promise<void> => {
    const sampleXml = await getSampleXml();
    const result = await ParseXml(sampleXml);
    if (typeof result === 'string') {
        console.error(result);
        return;
    }
    const origFormula = result.getFormula();
    const newFormula = origFormula.replace(
        '"This is an example."',
        '"This is the demonstration."'
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

RunDemo();

// */

export { ParseXml };
