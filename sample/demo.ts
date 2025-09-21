import {
    ExcelCustomXml,
    MashupFormulaSectionDefault,
    // ExcelZip,
} from '../dist/esm';

(async () => {
    const demoXml = require('./demo.json') as string;
    const excelXml = await ExcelCustomXml.create(demoXml);
    const { datamashup } = excelXml;
    const powerQuery = datamashup.rootItems.find((o) =>
        o.path.endsWith(MashupFormulaSectionDefault)
    );
    if (!powerQuery) {
        console.log('No Power Query found.');
        return;
    }
    datamashup.setFileContents(
        powerQuery,
        'section Section1;\n\nshared Test = let\r\n    result = #table(1, {{"This is an example that has had its query edited."}})\r\nin\r\n    result;'
    );
    const newXml = await excelXml.pack();
    if (!newXml) {
        console.log('Unable to pack the XML file.');
        return;
    }
    console.log('Summary datamashup demo:');
    console.log(`New size: ${newXml.length}`);
    console.log(`Old size: ${demoXml.length}`);
    console.log(`${demoXml.length - newXml.length} delta`);
})();

// (async () => {
//     const demoXlsx = new Uint8Array(); // load any Excel xlsx file that contains a PowerQuery
//     const excelZip = await ExcelZip.unzip(demoXlsx);
//     const powerQuery = await excelZip.getPowerQueryFile();
//     if (!powerQuery) {
//         console.log('No Power Query found.');
//         return;
//     }
//     await excelZip.setPowerQueryFile(
//         powerQuery,
//         'section Section1;\n\nshared Test = let\r\n    result = #table(1, {{"This is an example that has had its query edited."}})\r\nin\r\n    result;'
//     );
//     const newXlsx = await excelZip.zip();
//     if (!newXlsx.ok) {
//         console.log('Unable to pack the xlsx file.');
//         return;
//     }
//     console.log('Summary xslx demo:');
//     console.log(`New size: ${newXlsx.data.length}`);
//     console.log(`Old size: ${demoXlsx.length}`);
//     console.log(`${demoXlsx.length - newXlsx.data.length} delta`);
// })();
