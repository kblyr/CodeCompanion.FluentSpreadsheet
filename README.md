# CodeCompanion.FluentSpreadsheet
[![GitHub release (latest SemVer)](https://img.shields.io/github/v/release/kblyr/CodeCompanion.FluentSpreadsheet?color=white&logo=github)](https://github.com/kblyr/CodeCompanion.FluentSpreadsheet)
[![Nuget version (CodeCompanion.FluentSpreadsheet)](https://img.shields.io/nuget/v/CodeCompanion.FluentSpreadsheet?logo=nuget)](https://www.nuget.org/packages/CodeCompanion.FluentSpreadsheet)
[![GitHub](https://img.shields.io/github/license/kblyr/CodeCompanion.FluentSpreadsheet)](https://raw.githubusercontent.com/kblyr/CodeCompanion.FluentSpreadsheet/main/LICENSE)

When building systems for a client, there will always be a requirement to generate reports, and sometimes we need to export it into spreadsheets.

There are different libraries that generate spreadsheets like EPPlus, GemBox.Spreadsheet and Spire.XLS. But the code you write to generate spreadsheet using EPPlus will be different when you use GemBox.Spreadsheet or Spire.XLS.

The purpose of this library is to give developers an API where they can switch and use different spreadsheet libraries without rewritten the whole code. And plus! We use the *Fluent API* coding pattern here, so you can chain methods.

## Samples
In this section, I will show some of the functionalities of the library.
```csharp
// Assuming you are using EPPlus

var factory = new FluentSpreadsheetFactory().New() // Create new workbook
    .AddWorksheet("Sheet 1") // Add new sheet with name = 'Sheet 1'
        .OnRange("A1") // Target the cell 'A1' in sheet 'Sheet 1'
            .WithValue(10) // Set the value of cell 'A1' to number 10
            .WithBorder(SpreadsheetBoxSide.All, SpreadsheetBorderStyle.Thin) // Surrounds the cell with thin border
        .OnRange(0, 1) // Target the cell 'B1' in sheet 'Sheet 1'; we use zero-based indexing
            .WithValue("ABC") // Set the value of cell 'B1' with string ABC
            .WithHorizontalAlignment(SpreadsheetHorizontalAlignment.Center) // Center-align the content of the cell horizontally
    .AddWorksheet("Sheet 2") // Add new sheet with name = 'Sheet 2'
        .OnRange(0, 0, 2, 4) // Target cell range 'A1:E3' in sheet 'Sheet 2'
            .Merge() // Merge the cell range
            .WithValue("This is a merged cell") // Set the value of the merged cell range
    .CopyWorksheet("Sheet 1", "Sheet 3") // Copy 'Sheet 1' and add new sheet with name = 'Sheet 3', all the content will be copied
    .DeleteWorksheet(0) // Delete sheet with index = 0 (Sheet 1)
    .SaveAs(@"C:\Spreadsheets\Sample.xlsx") // Save the workbook into a .xlsx file
    ;
```
