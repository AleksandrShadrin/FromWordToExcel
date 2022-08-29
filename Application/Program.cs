using Application.Services.ExcelService;
using Application.Services.ValueObjects;
using Application.Services.WordService;


var documentPath = Path.GetFullPath(Path.Combine("."));
string firstLine = "";

using (var wordService = WordService.OpenWordFile(Path.Combine(documentPath, "test.docx")))
{
    firstLine = wordService.ReadLines().FirstOrDefault();
}

using (var excelService = ExcelService.OpenExcelFile(Path.Combine(documentPath, "test.xlsx")))
{
    var b2 = new ExcelCellCoords(2, 2);
    excelService.WriteToCell(b2, firstLine);
}

