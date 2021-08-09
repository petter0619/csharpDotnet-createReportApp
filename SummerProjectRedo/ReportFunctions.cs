using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Net;
using System.Threading.Tasks;
using Microsoft.Azure.Functions.Worker;
using Microsoft.Azure.Functions.Worker.Http;
using Microsoft.Extensions.Logging;
using System.Web;
using System.Text;
using System.Net.Http;
using System.Text.Json;
// EPPlus (Excel)
using OfficeOpenXml;
using OfficeOpenXml.Style;
// HTML to PDF
using Aspose.Pdf;
// Package for getting mimetypes
using MimeTypes;


namespace CreateAndDownloadExcelReport
{
    public static class ReportFunctions
    {
        [Function("CreateReport")]
        public static async Task<dynamic> CreateReport(
            [HttpTrigger(AuthorizationLevel.Anonymous, "get", Route = "Project/{projectId}/report")] HttpRequestData req,
            string projectId,
            FunctionContext executionContext
        )
        {
            // --------- Instantiate logger ---------
            var logger = executionContext.GetLogger("CreateReport");
            logger.LogInformation($"Processing request for creation and export of report for project {projectId}");

            // --------- Parse desired report format from querystring --------
            var queryParams = HttpUtility.ParseQueryString(req.Url.Query);
            string reportFormat;

            if (queryParams.HasKeys() && queryParams.Get("format") != null)
            {
                reportFormat = queryParams.Get("format");
            }
            else
            {
                var errorResponse = await ErrorResponse(
                    req.CreateResponse(HttpStatusCode.BadRequest),
                    new Exception("Please specify a file format (pdf || xlsx) via the 'format' query parameter.")
                );
                return errorResponse;
            }

            // --------- Generate Report ---------
            var fileExtension = reportFormat.ToLower() == "excel"
                ? "xlsx"
                : reportFormat;

            string reportName = queryParams.Get("name") == null
                ? $"ProjektRapport-{DateTime.Now.ToString("yyyy'-'MM'-'dd")}.{fileExtension}"
                : $"{queryParams.Get("name").Replace(" ", "")}.{fileExtension}";

            (byte[] Report, Exception Error) fileData = (null, null);

            /* DIAGNOSTIC INFO */ logger.LogInformation($"----- Generating new {fileExtension} project report -----");
            /* DIAGNOSTIC INFO */ var watch = System.Diagnostics.Stopwatch.StartNew();

            switch (reportFormat.ToLower())
            {
                case "excel":
                    fileData = await GenerateNewExcelReportFromTemplate(fileData);
                    break;
                case "xlsx":
                    fileData = await GenerateNewExcelReportFromTemplate(fileData);
                    break;
                case "pdf":
                    fileData = await GenerateNewPdfReportFromTemplate(fileData);
                    break;
                default:
                    var errorResponse = await ErrorResponse(
                        req.CreateResponse(HttpStatusCode.BadRequest),
                        new Exception($"Format not supported: {reportFormat}")
                    );
                    return errorResponse;
            }

            /* DIAGNOSTIC INFO */ watch.Stop();
            /* DIAGNOSTIC INFO */ logger.LogInformation($"----- {fileExtension} report created in {watch.ElapsedMilliseconds} ms -----");

            // --------- Response ---------

            if (fileData.Error != null)
            {
                logger.LogError("ERROR during report creation: \n" + fileData.Error.InnerException + "\n" + fileData.Error.Message);

                var errorResponse = await ErrorResponse(
                    req.CreateResponse(HttpStatusCode.InternalServerError),
                    fileData.Error
                );
                return errorResponse;
            }
            else
            {
                var successResponse = await FileDownloadResponse(
                    req,
                    reportName,
                    fileData.Report
                );
                return successResponse;
            }
        }

        public static async Task<(byte[] Report, Exception Error)> GenerateNewExcelReportFromTemplate(
            (byte[] Report, Exception Error) output
        )
        {
            try
            {
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

                var excelTemplate = File.OpenRead(Path.Combine(Directory.GetCurrentDirectory(), "ExcelTemplate.xlsx"));

                using (var package = new ExcelPackage(excelTemplate))
                {

                    /* ---------- Input report custom data ---------- */

                    // WORKSHEET: Projektinformation
                    var ws_projectInformation = package.Workbook.Worksheets["ProjektInformation"];

                    ws_projectInformation.Cells["B5"].Value = "AFRY Head Office"; // Projektnamn
                    ws_projectInformation.Cells["B9"].Value = "99435"; // Projektkod
                    ws_projectInformation.Cells["E12"].Value = 123456; // Bruttoarea
                    ws_projectInformation.Cells["E14"].Value = 12; // Antal Våningar
                    ws_projectInformation.Cells["E16"].Value = 1234; // Byggnadsarea

                    // WORKSHEET: TotalKostnad
                    var ws_totalCost = package.Workbook.Worksheets["TotalKostnad"];

                    ws_totalCost.Cells["D5"].Value = "AFRY Head Office"; // Projektnamn
                    ws_totalCost.Cells["L4"].Value = "99435"; // Projektkod
                    ws_totalCost.Cells["L5"].Value = DateTime.Now.ToString("yyyy'-'MM'-'dd"); // Datum
                    ws_totalCost.Cells["L6"].Value = 123456; // Bruttoarea

                    ws_totalCost.Cells["J11"].Value = 5000000; // Mängd
                    ws_totalCost.Cells["J13"].Value = 0; // EnH
                    ws_totalCost.Cells["J15"].Value = 3000000; // Material
                    ws_totalCost.Cells["J17"].Value = 2000000; // Arbete
                    ws_totalCost.Cells["J19"].Value = 1000000; // Maskin
                    ws_totalCost.Cells["J21"].Value = 0; // UE
                    ws_totalCost.Cells["J23"].Value = 0; // Pris
                    ws_totalCost.Cells["J31"].Value = 30000000; // TOTAL

                    ws_totalCost.Cells["L11"].Value = 5000000 / 123456; // Mängd per m2
                    ws_totalCost.Cells["L13"].Value = 0; // EnH per m2
                    ws_totalCost.Cells["L15"].Value = 3000000 / 123456; // Material per m2
                    ws_totalCost.Cells["L17"].Value = 2000000 / 123456; // Arbete per m2
                    ws_totalCost.Cells["L19"].Value = 1000000 / 123456; // Maskin per m2
                    ws_totalCost.Cells["L21"].Value = 0; // UE per m2
                    ws_totalCost.Cells["L23"].Value = 0; // Pris per m2
                    ws_totalCost.Cells["L31"].Value = 30000000 / 123456; // TOTAL per m2

                    // WORKSHEET: Substrukturer
                    var ws_subStructures = package.Workbook.Worksheets["SubStrukturer"];

                    ws_subStructures.Cells["D5"].Value = "AFRY Head Office"; // Projektnamn
                    ws_subStructures.Cells["O5"].Value = DateTime.Now.ToString("yyyy'-'MM'-'dd"); // Datum

                    // SUBSTRUCTURE MOCK DATA
                    var substructureMock = new[] {
                        new {
                            Name = "Garage",
                            Mangd = 321,
                            Enh = 0,
                            Material = 322,
                            Arbete = 323,
                            Maskin = 324,
                            Ue = 0,
                            Pris = 0,
                            Total = 1234
                        },
                        new
                        {
                            Name = "Basement",
                            Mangd = 321,
                            Enh = 0,
                            Material = 322,
                            Arbete = 323,
                            Maskin = 324,
                            Ue = 0,
                            Pris = 0,
                            Total = 1234
                        },
                        new
                        {
                            Name = "Attic",
                            Mangd = 321,
                            Enh = 0,
                            Material = 322,
                            Arbete = 323,
                            Maskin = 324,
                            Ue = 0,
                            Pris = 0,
                            Total = 1234
                        }
                    };

                    int startingRow = 11;
                    System.Drawing.Color colFromHex = ColorTranslator.FromHtml("#FFFFCC");

                    foreach (var s in substructureMock)
                    {
                        // Row
                        ws_subStructures.Row(startingRow).Style.Font.Bold = true;
                        ws_subStructures.Row(startingRow).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        ws_subStructures.Column(2).AutoFit();

                        ws_subStructures.Cells[$"B{startingRow + 1}:R{startingRow + 1}"].Style.Border.Bottom.Style = ExcelBorderStyle.Dotted;

                        // SubStructure name
                        var nameCell = ws_subStructures.Cells[$"B{startingRow}"];
                        nameCell.Value = s.Name;
                        nameCell.Style.Font.Size = 12;
                        nameCell.Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;

                        // Mängd
                        var mangdCell = ws_subStructures.Cells[$"D{startingRow}"];
                        mangdCell.Value = s.Mangd;
                        mangdCell.Style.Fill.PatternType = ExcelFillStyle.Solid;
                        mangdCell.Style.Fill.BackgroundColor.SetColor(colFromHex);
                        mangdCell.Style.Border.Top.Style = ExcelBorderStyle.Thin;
                        mangdCell.Style.Border.Right.Style = ExcelBorderStyle.Thin;
                        mangdCell.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                        mangdCell.Style.Border.Left.Style = ExcelBorderStyle.Thin;


                        // EnH
                        var enhCell = ws_subStructures.Cells[$"F{startingRow}"];
                        enhCell.Value = s.Enh;
                        enhCell.Style.Fill.PatternType = ExcelFillStyle.Solid;
                        enhCell.Style.Fill.BackgroundColor.SetColor(colFromHex);
                        enhCell.Style.Border.Top.Style = ExcelBorderStyle.Thin;
                        enhCell.Style.Border.Right.Style = ExcelBorderStyle.Thin;
                        enhCell.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                        enhCell.Style.Border.Left.Style = ExcelBorderStyle.Thin;

                        // Material
                        var materialCell = ws_subStructures.Cells[$"H{startingRow}"];
                        materialCell.Value = s.Material;
                        materialCell.Style.Fill.PatternType = ExcelFillStyle.Solid;
                        materialCell.Style.Fill.BackgroundColor.SetColor(colFromHex);
                        materialCell.Style.Border.Top.Style = ExcelBorderStyle.Thin;
                        materialCell.Style.Border.Right.Style = ExcelBorderStyle.Thin;
                        materialCell.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                        materialCell.Style.Border.Left.Style = ExcelBorderStyle.Thin;

                        // Arbete
                        var arbeteCell = ws_subStructures.Cells[$"J{startingRow}"];
                        arbeteCell.Value = s.Arbete;
                        arbeteCell.Style.Fill.PatternType = ExcelFillStyle.Solid;
                        arbeteCell.Style.Fill.BackgroundColor.SetColor(colFromHex);
                        arbeteCell.Style.Border.Top.Style = ExcelBorderStyle.Thin;
                        arbeteCell.Style.Border.Right.Style = ExcelBorderStyle.Thin;
                        arbeteCell.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                        arbeteCell.Style.Border.Left.Style = ExcelBorderStyle.Thin;

                        // Maskin
                        var maskinCell = ws_subStructures.Cells[$"L{startingRow}"];
                        maskinCell.Value = s.Maskin;
                        maskinCell.Style.Fill.PatternType = ExcelFillStyle.Solid;
                        maskinCell.Style.Fill.BackgroundColor.SetColor(colFromHex);
                        maskinCell.Style.Border.Top.Style = ExcelBorderStyle.Thin;
                        maskinCell.Style.Border.Right.Style = ExcelBorderStyle.Thin;
                        maskinCell.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                        maskinCell.Style.Border.Left.Style = ExcelBorderStyle.Thin;

                        // UE
                        var ueCell = ws_subStructures.Cells[$"N{startingRow}"];
                        ueCell.Value = s.Ue;
                        ueCell.Style.Fill.PatternType = ExcelFillStyle.Solid;
                        ueCell.Style.Fill.BackgroundColor.SetColor(colFromHex);
                        ueCell.Style.Border.Top.Style = ExcelBorderStyle.Thin;
                        ueCell.Style.Border.Right.Style = ExcelBorderStyle.Thin;
                        ueCell.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                        ueCell.Style.Border.Left.Style = ExcelBorderStyle.Thin;

                        // Pris
                        var prisCell = ws_subStructures.Cells[$"P{startingRow}"];
                        prisCell.Value = s.Pris;
                        prisCell.Style.Fill.PatternType = ExcelFillStyle.Solid;
                        prisCell.Style.Fill.BackgroundColor.SetColor(colFromHex);
                        prisCell.Style.Border.Top.Style = ExcelBorderStyle.Thin;
                        prisCell.Style.Border.Right.Style = ExcelBorderStyle.Thin;
                        prisCell.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                        prisCell.Style.Border.Left.Style = ExcelBorderStyle.Thin;

                        // TOTAL
                        var totalCell = ws_subStructures.Cells[$"R{startingRow}"];
                        totalCell.Value = s.Total;
                        totalCell.Style.Fill.PatternType = ExcelFillStyle.Solid;
                        totalCell.Style.Fill.BackgroundColor.SetColor(colFromHex);
                        totalCell.Style.Border.Top.Style = ExcelBorderStyle.Double;
                        totalCell.Style.Border.Right.Style = ExcelBorderStyle.Double;
                        totalCell.Style.Border.Bottom.Style = ExcelBorderStyle.Double;
                        totalCell.Style.Border.Left.Style = ExcelBorderStyle.Double;

                        // Increment rowcount
                        startingRow += 3;
                    }

                    /* ---------- Return byte[] for direct file download ---------- */
                    output.Report = await package.GetAsByteArrayAsync();
                    return output;
                }
            }
            catch (Exception ex)
            {
                output.Error = ex;
                return output;
            }
        }

        public static async Task<(byte[] Report, Exception Error)> GenerateNewPdfReportFromTemplate(
            (byte[] Report, Exception Error) output
        )
        {
            try { 
                var templatePath = Path.Combine(Directory.GetCurrentDirectory(), "PDFTemplate.html");
                string placeholderHtml = await File.ReadAllTextAsync(templatePath);

                // --------- Replace {{{placeholders}}} in htmlString with data ---------

                string projectInfoData = placeholderHtml
                    .Replace("{{{projectName}}}", "AFRY Head Office")       // Projektnamn
                    .Replace("{{{projectCode}}}", "99435")                  // Projektkod
                    .Replace("{{{bruttoarea}}}", "123456")                  // Bruttoarea
                    .Replace("{{{antalVaningar}}}", "12")                   // Antal Våningar
                    .Replace("{{{byggnadsarea}}}", "1234");                 // Byggnadsarea

                string productionCostData = projectInfoData
                    .Replace("{{{date}}}", DateTime.Now.ToString("yyyy'-'MM'-'dd")) // Date
                    .Replace("{{{totaltMangd}}}", "5000000")                        // Mängd
                    .Replace("{{{totaltEnh}}}", "0")                                // EnH
                    .Replace("{{{totaltMaterial}}}", "3000000")                     // Material
                    .Replace("{{{totaltArbete}}}", "2000000")                       // Arbete
                    .Replace("{{{totaltMaskin}}}", "1000000")                       // Maskin
                    .Replace("{{{totaltUe}}}", "0")                                 // UE
                    .Replace("{{{totaltPris}}}", "0")                               // Pris
                    .Replace("{{{totaltTotal}}}", "30000000")                       // TOTAL
                    .Replace("{{{m2Mangd}}}", (5000000/123456).ToString())          // Mängd per m2
                    .Replace("{{{m2Enh}}}", "0")                                    // EnH per m2
                    .Replace("{{{m2Material}}}", (3000000 / 123456).ToString())     // Material per m2
                    .Replace("{{{m2Arbete}}}", (2000000 / 123456).ToString())       // Arbete per m2
                    .Replace("{{{m2Maskin}}}", (1000000 / 123456).ToString())       // Maskin per m2
                    .Replace("{{{m2Ue}}}", "0")                                     // UE per m2 
                    .Replace("{{{m2Pris}}}", "0")                                   // Pris per m2
                    .Replace("{{{m2Total}}}", (30000000 / 123456).ToString());      // TOTAL per m2

                // SUBSTRUCTURE MOCK DATA
                var substructureMock = new[] {
                    new {
                            Name = "Garage",
                            Mangd = 321,
                            Enh = 0,
                            Material = 322,
                            Arbete = 323,
                            Maskin = 324,
                            Ue = 0,
                            Pris = 0,
                            Total = 1234
                    },
                    new {
                            Name = "Basement",
                            Mangd = 321,
                            Enh = 0,
                            Material = 322,
                            Arbete = 323,
                            Maskin = 324,
                            Ue = 0,
                            Pris = 0,
                            Total = 1234
                    }
                };

                string substructureRows = "";
            
                foreach (var s in substructureMock)
                {
                    var tableHtml = await File.ReadAllTextAsync(Path.Combine(Directory.GetCurrentDirectory(), "SubstructureTableTemplate.html"));
                    substructureRows += tableHtml
                        .Replace("{{{substructureName}}}", s.Name)
                        .Replace("{{{substructureMangd}}}", s.Mangd.ToString())
                        .Replace("{{{substructureEnh}}}", s.Enh.ToString())
                        .Replace("{{{substructureMaterial}}}", s.Material.ToString())
                        .Replace("{{{substructureArbete}}}", s.Arbete.ToString())
                        .Replace("{{{substructureMaskin}}}", s.Maskin.ToString())
                        .Replace("{{{substructureUe}}}", s.Ue.ToString())
                        .Replace("{{{substructurePris}}}", s.Pris.ToString())
                        .Replace("{{{substructureTotal}}}", s.Total.ToString());
                }
            
                string dataHtml = productionCostData.Replace("{{{substructureRows}}}", substructureRows);

                // --------- Convert htmlString into PDF byte[] ---------
                /*
                // via SautinSoft.PdfMetamorphosis
                SautinSoft.PdfMetamorphosis p = new SautinSoft.PdfMetamorphosis();
                var pdfBytes = p.HtmlToPdfConvertStringToByte(dataHtml);
                */

                // via Aspose.Net
                byte[] pdfByteArray;

                HtmlLoadOptions objLoadOptions = new HtmlLoadOptions();
                Document doc = new Document(new MemoryStream(Encoding.UTF8.GetBytes(dataHtml)), objLoadOptions);
                PageCollection pageCollection = doc.Pages;

                foreach (var p in pageCollection)
                {
                    p.SetPageSize(657.6, 842.4);
                }

                using (MemoryStream ms = new MemoryStream())
                {
                    doc.Save(ms);
                    pdfByteArray = ms.ToArray();
                }

                // --------- Send output --------
                output.Report = pdfByteArray;
                return output;
            }
            catch (Exception ex) {
                output.Error = ex;
                return output;
            }
        }

        public static async Task<HttpResponseData> FileDownloadResponse(
            HttpRequestData request,
            string fileNameWithExtension,
            byte[] fileAsByteArray
        )
        {
            string extension = Path.GetExtension(fileNameWithExtension);
            var response = request.CreateResponse(HttpStatusCode.OK);
            response.Headers.Add("Content-Disposition", $"attachment;filename={fileNameWithExtension}");
            response.Headers.Add("Content-Type", $"{MimeTypeMap.GetMimeType(extension)}");
            await response.Body.WriteAsync(fileAsByteArray);
            return response;
        }

        public static async Task<HttpResponseData> ErrorResponse(
            HttpResponseData response,
            Exception ex = null
        )
        {
            if (ex != null)
            {
                string jsonString = JsonSerializer.Serialize(new
                {
                    StatusCode = response.StatusCode,
                    StatusName = response.StatusCode.ToString(),
                    Message = ex.Message
                });
                await response.WriteStringAsync(jsonString);
            }
            return response;
        }
    }
}
