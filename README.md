# excel-datamashup

This sample project contains code to convert a Excel customXml item1.xml file into a usable data structure.

The various Excel formats `xlsx`, `xlsm`, `xlsb` are ZIP based and thus can be extracted.

The contents will in situations where Power Query is used, contain a customXml folder with a item1.xml file that contains the relevant data structure in binary format.

This binary format can be processed into something we can read, edit, then re-package back into binary format.

The goal of this project is to faciliate processing a Excel file, then being able to edit and save it in both browser and terminal modes.

## API

<details>
  <summary>Directly editing `customXml\item1.xml` (DataMashup) using `ExcelCustomXml`.</summary>

```ts
import { type UnzippedItem, ExcelCustomXml } from 'excel-datamashup';

const xml: string = '...';

// returns a working instance of the class
const excelXml: ExcelCustomXml = await ExcelCustomXml.create(xml);

// find the power query file
const powerQuery: UnzippedItem | undefined = excelXml.datamashup.rootItems.find(
    (o) => o.path.endsWith('Section1.m')
);

// if found, set its contents to something else
if (powerQuery) {
    excelXml.datamashup.setFileContents(powerQuery, '...');

    // always reset permissions when editing
    await excelXml.datamashup.resetPermissions();
}

// pack the data back to a xml string, then write it back to the `customXml\item1.xml` file using your favorite zip editing library
const newXml: string | undefined = await excelXml.pack();
```

</details>

<details>
  <summary>Editing an Excel xlsx file using `ExcelZip`.</summary>

```ts
import { type Result, ExcelZip, UnzippedItem } from 'excel-datamashup';

// read and store the binary zip data as number array or Uint8Array
const zip = new Uint8Array();

// process the zip into a more manageable object
const excelZip: ExcelZip = await ExcelZip.unzip(zip);

// get the power query contents
const powerQuery: UnzippedItem | undefined = await excelZip.getPowerQueryFile();

// modify the power query contents
if (powerQuery) {
    await excelZip.setPowerQueryFile(
        powerQuery,
        'section Section1;\n\nshared Test = let\r\n    result = #table(1, {{"This is an example."}})\r\nin\r\n    result;'
    );
}

// zip the contents back to an Excel file
const result: Result<Uint8Array> = await excelZip.zip();

// evaluate if it was successfull
if (result.ok) {
    console.log('Save the xlsx file:', result.data.length);
} else {
    console.log('Unable to create xlsx file.');
}
```

</details>

## Sample

The sample folder contain an example file that contains a Power Query, along with some additional sample code.

## Resources

-   https://bengribaudo.com/blog/2020/04/22/5198/data-mashup-binary-stream
-   https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-qdeff/27b1dd1e-7de8-45d9-9c84-dfcc7a802e37
