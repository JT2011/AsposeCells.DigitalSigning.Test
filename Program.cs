// See https://aka.ms/new-console-template for more information

using Aspose.Cells;
using AsposeCells.DigitalSigning.Test;

Console.WriteLine("Hello, World!");

//TODO:Upload file and output signed file
//TODO:Utilize Aspose cells workbook.VbaProject.Sign with pfx cert file

var workbook = WorkbookExtensions.CreateWorkbook();
var digitalSignature = WorkbookExtensions.GetDigitalSignature();
workbook.VbaProject.Sign(digitalSignature);
var stream = new MemoryStream();
workbook.Save(stream, SaveFormat.Xlsm);
MyStream.Main(stream.ToArray());
Console.WriteLine($"Is file properly signed: {workbook.VbaProject.IsValidSigned}");