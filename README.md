# excel-datamashup

This sample project contains code to convert a Excel customXml item1.xml file into a usable data structure.

The various Excel formats `xlsx`, `xlsm`, `xlsb` are ZIP based and thus can be extracted.

The contents will in situations where Power Query is used, contain a customXml folder with a item1.xml file that contains the relevant data structure in binary format.

This binary format can be processed into something we can read, edit, then re-package back into binary format.

The goal of this project is to faciliate processing a Excel file, then being able to edit and save it in both browser and node.

## API

```ts
import { ParseXml } from 'excel-datamashup';

// extract the contents of `customXml\item1.xml` using your favorite zip editing library
const xml = `...`;

// returns `ParseResult` object or a string corresponding to the `ParseError` type
const result = await ParseXml(xml);

// the object has some helper methods along with the raw data for manipulation
const originalFormula = result.getFormula();

// edit the original or create a entirely new power query
const newFormula = `...`;

// replace the formula with your new one
result.setFormula(newFormula);

// always reset permissions when editing
result.resetPermissions();

// update the contents of `customXml\item1.xml` by writing this new content using your favorite zip editing library
const newXml = await result.save();
```

## Sample

The sample folder contain an example. It contains a Power Query that simply outputs a table with some text.

If you want to test using your own Excel file you can use the following PowerShell snippet:

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

## Resources

- https://bengribaudo.com/blog/2020/04/22/5198/data-mashup-binary-stream
- https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-qdeff/27b1dd1e-7de8-45d9-9c84-dfcc7a802e37
