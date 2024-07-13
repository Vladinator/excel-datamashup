# excel-datamashup

This sample project contains code to convert a Excel customXml item1.xml file into a usable data structure.

The various Excel formats `xlsx`, `xlsm`, `xlsb` are ZIP based and thus can be extracted.

The contents will in situations where Power Query is used, contain a customXml folder with a item1.xml file that contains the relevant data structure in binary format.

This binary format can be processed into something we can read, edit, then re-package back into binary format.

The goal of this project is to faciliate processing a Excel file, then being able to edit and save it in both browser and terminal modes.

## API

You can work with the library in either ExcelZip mode which provides a certain level of wrapper for you.

```ts
import { type UnzippedExcel, ExcelZip } from 'excel-datamashup';

// read and store the binary zip data as number array, Uint8Array or Buffer
const zip = new Uint8Array();

// process the zip into a more manageable object
const excelZip: UnzippedExcel = await ExcelZip(zip);

// the object has some helper methods along with the raw data for manipulation
const originalFormula: string = excelZip.getFormula();

// modify or replace the formula entirely
const newFormula: string = originalFormula.replace('"Some Text"', '"New Text"');

// replace the formula with your new one
// this method will also call the internal `excelZip.datamashup.result.resetPermissions()` method for you
excelZip.setFormula(newFormula);

// zip the contents back to an Excel file
const newZip: Buffer = await excelZip.save();
```

The more direct approach is to simply focus on processing the DataMashup XML file directly.

```ts
import { ParseXml } from 'excel-datamashup';

// extract the contents of `customXml\item1.xml` using your favorite zip editing library
const xml: string = `...`;

// returns `ParseResult` object or a string corresponding to the `ParseError` type
const result: ParseResult = await ParseXml(xml);

// the object has some helper methods along with the raw data for manipulation
const originalFormula: string = result.getFormula();

// edit the original or create a entirely new power query
const newFormula: string = `...`;

// replace the formula with your new one
result.setFormula(newFormula);

// always reset permissions when editing
result.resetPermissions();

// update the contents of `customXml\item1.xml` by writing this new content using your favorite zip editing library
const newXml: string = await result.save();
```

## Sample

The sample folder contain an example file that contains a Power Query. It simply outputs text to a table.

If you have this project checked out, simply uncomment the `src\index.ts` line 10 to enable the demo behavior.

Build the project and run `dist\index.js` in a browser, it will ask for you to upload a Excel file, which it will edit then download.

<details>
  <summary>If you want to test using your own Excel file you can use the following PowerShell snippet to have it update `sample\excel.json` with your own Power Query.</summary>

```pwsh
$inputFile = ".\sample\demo.xlsb"
$outputFile = ".\sample\demo.json"
$zipName = "customXml/item1.xml"
$tempFile = [System.IO.Path]::GetTempFileName()
Add-Type -AssemblyName System.IO.Compression.FileSystem
$zip = [System.IO.Compression.ZipFile]::OpenRead($inputFile)
try {
    $entry = $zip.Entries | Where-Object { $_.FullName -eq $zipName }
    $stream = $entry.Open()
    try {
        $reader = [IO.StreamReader]::new($stream)
        $writer = [IO.StreamWriter]::new($tempFile)
        try {
            while (-not $reader.EndOfStream) {
                $line = $reader.ReadLine()
                $writer.WriteLine($line)
            }
        } finally {
            $writer.Close()
            $reader.Close()
        }
    } finally {
        $stream.Close()
    }
} finally {
    $zip.Dispose()
}
(Get-Content -Path $tempFile) -join "" | ConvertTo-Json -Compress | Set-Content -Path $outputFile
Remove-Item -Path $tempFile
```
</details>

## Resources

-   https://bengribaudo.com/blog/2020/04/22/5198/data-mashup-binary-stream
-   https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-qdeff/27b1dd1e-7de8-45d9-9c84-dfcc7a802e37
