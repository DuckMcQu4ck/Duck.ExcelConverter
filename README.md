# ExcelConverter

`ExcelConverter` is a .NET library designed to convert Excel data into a list of strongly-typed objects. It uses NPOI to handle Excel files and offers asynchronous methods for conversion.

## Features

- Convert Excel data to a list of custom objects asynchronously.
- Supports both stream and byte array inputs.
- Flexible column mapping to map Excel columns to object properties.
- Handles different cell types (Boolean, Numeric, String, Formula).

## Installation

Ensure you have the following dependencies in your project:

- `NPOI`

```bash
dotnet add package NPOI --version 2.7.1

```

## Usage example

```csharp
public class YourObject
{
    public string ObjectProperty1 { get; set; }
    public int ObjectProperty2 { get; set; }
}    

var converter = new ExcelConverter();
var columnMappings = new Dictionary<string, string>
            {
                {"ExcelColumn1", "ObjectProperty1"},
                {"ExcelColumn2", "ObjectProperty2"}
            };

using (var fileStream = File.OpenRead("path/to/your/excel/file.xlsx"))
{
    var result = await converter.ConvertToListAsync<YourObject>(fileStream, columnMappings);
}
```
