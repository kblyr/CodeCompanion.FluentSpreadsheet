using CodeCompanion.FluentSpreadsheet;
using CodeCompanion.FluentSpreadsheet.Formulas;
using CodeCompanion.FluentSpreadsheet.Formulas.Arguments;
using CodeCompanion.FluentSpreadsheet.Styles;
using OfficeOpenXml;

ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

var factory = new FluentSpreadsheetFactory();
var borderDemo_itemCounter = 0;
factory.New()
    .AddWorksheet("Border Demo")
        .OnRange(1, 1, 1, 5)
        .Merge()
        .WithValue("Header")
        .WithFontBold()
        .WithBorder(SpreadsheetBoxSide.All, SpreadsheetBorderStyle.Thin)
        .WithHorizontalAlignment(SpreadsheetHorizontalAlignment.Center)
        .Worksheet
        .While(() => borderDemo_itemCounter < 10, worksheet => {
            worksheet
                .OnRange(2 + borderDemo_itemCounter, 1)
                    .WithValue(borderDemo_itemCounter + 1)
                    .WithBorder(SpreadsheetBoxSide.Bottom | SpreadsheetBoxSide.LeftRight, SpreadsheetBorderStyle.Thin)
                .OnRange(2 + borderDemo_itemCounter, 2)
                    .FxAverage(RangeRef.From(worksheet.OnRange(2, 1, 2 + 9, 1)));
            borderDemo_itemCounter++;
        })

    .SaveAs(Path.Combine(Directory.GetCurrentDirectory(), "test.xlsx"));