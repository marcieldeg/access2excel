# Description
Access2Excel convert a MSAccess database into a Excel spreadsheet. It's a pure java implementation, using POI (https://poi.apache.org/ ) and Jackcess (http://jackcess.sourceforge.net/).

## How to use:
`java -jar Access2Excel.jar -inputFile=<inputFile> [-outputFile=<outputFile>] [-format=<format>]`

where

`<inputFile>`: A file name to Access database (.MDB or .ACCDB)<br/>
`<outputFile>`: A file name to Excel streadsheet (optional, default "inputFile.xls" or "inputFile.xlsx")<br/>
`<format>`: Output format ("XLS" or "XLSX", optional, default "XLSX")
