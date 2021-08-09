using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Net;
using System.Net.Http;
using System.Threading.Tasks;
using Azure.Storage.Blobs;
using Microsoft.Azure.Functions.Worker;
using Microsoft.Azure.Functions.Worker.Http;
using Microsoft.Extensions.Logging;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System.Web;
using System.Text.Json;
using Azure.Storage.Sas;
using System.Threading;
using System.Collections.Concurrent;

using System.Linq;

namespace CreateAndDownloadExcelReport
{
    public static class StorageReportFunctions
    {
        [Function("CreateReportFromStorage")]
        public static async Task<HttpResponseData> CreateReportFromStorage(
            [HttpTrigger(AuthorizationLevel.Function, "get", Route = "Project/{projectId}/storage-report")] HttpRequestData req,
            string projectId,
            FunctionContext executionContext
        )
        {
            // --------- Logger setup ---------
            var logger = executionContext.GetLogger("CreateReportFromStorage");
            logger.LogInformation($"Requesting the creation and export of an Excel report for project {projectId}");

            // --------- Get format from querystring --------
            var queryParams = HttpUtility.ParseQueryString(req.Url.Query);
            string reportFormat;

            if (queryParams.HasKeys() && queryParams.Get("format") != null)
            {
                reportFormat = queryParams.Get("format");
            }
            else
            {
                return req.CreateResponse(HttpStatusCode.BadRequest);
            }

            // --------- Generate Excel Report ---------
            string reportName = queryParams.Get("name") == null
                ? $"ProjektRapport-{DateTime.Now.ToString("yyyy'-'MM'-'dd")}"
                : queryParams.Get("name");

            string conStr = Environment.GetEnvironmentVariable("AzureBlobStorageConnectionString", EnvironmentVariableTarget.Process);
            string templateContainer = "report-templates";
            var containerClient = new BlobContainerClient(conStr, templateContainer);

            ((string Name, string Uri) Report, Exception Error ) output = ((null, null), null);

            switch (reportFormat.ToLower())
            {
                case "excel":
                    output = await GenerateNewExcelReportFromTemplate(containerClient, "ExcelTemplate.xlsx", reportName, output);
                    break;
                default:
                    return req.CreateResponse(HttpStatusCode.BadRequest);
            }

            // --------- Upload File to Azure Storage and send SAS link ---------

            HttpResponseData response;
            if (output.Error != null)
            {
                response = req.CreateResponse(HttpStatusCode.InternalServerError);
                response.Headers.Add("Content-Type", "application/json");
                var resJson = new
                {
                    Error = output.Error,
                };
                await response.WriteStringAsync(JsonSerializer.Serialize(resJson));
            }
            else
            {
                response = req.CreateResponse(HttpStatusCode.Created);
                response.Headers.Add("Content-Type", "application/json");
                var resJson = new
                {
                    Name = output.Report.Name,
                    Uri = output.Report.Uri
                };
                await response.WriteStringAsync(JsonSerializer.Serialize(resJson));
            }
            return response;
        }

        public static async Task<((string Name, string Uri) Report, Exception Error)> GenerateNewExcelReportFromTemplate(
            BlobContainerClient container,
            string templateToUse,
            string reportName,
            ((string Name, string Uri) Report, Exception Error) output
        )
        {
            try
            {
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

                var blockBlob = container.GetBlobClient(templateToUse);

                using (var memorystream = new MemoryStream())
                {
                await blockBlob.DownloadToAsync(memorystream);
                memorystream.Position = 0;

                    using (var package = new ExcelPackage(memorystream))
                    using (var fileStream = new MemoryStream())
                    {

                        /* ---------- Input report custom data ---------- */

                        // Projektinformation
                        var ws_projectInformation = package.Workbook.Worksheets["ProjektInformation"];

                        ws_projectInformation.Cells["B5"].Value = "AFRY Head Office"; // Projektnamn
                        ws_projectInformation.Cells["B9"].Value = "99435"; // Projektkod
                        ws_projectInformation.Cells["E12"].Value = 123456; // Bruttoarea
                        ws_projectInformation.Cells["E14"].Value = 12; // Antal Våningar
                        ws_projectInformation.Cells["E16"].Value = 1234; // Byggnadsarea

                        // TotalKostnad
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

                        // TotalKostnad
                        var ws_subStructures = package.Workbook.Worksheets["SubStrukturer"];

                        ws_subStructures.Cells["D5"].Value = "AFRY Head Office"; // Projektnamn
                        ws_subStructures.Cells["O5"].Value = DateTime.Now.ToString("yyyy'-'MM'-'dd"); // Datum

                        // MOCK DATA
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
                        Color colFromHex = ColorTranslator.FromHtml("#FFFFCC");

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

                        /* ---------- Add ExcelPackage to stream ---------- */
                        await package.SaveAsAsync(fileStream);
                        fileStream.Position = 0;

                        /* ---------- Upload to Blob Storage ---------- */
                        var newBlockBlob = container.GetBlobClient(reportName + Guid.NewGuid() + ".xlsx");
                        await newBlockBlob.UploadAsync(fileStream);

                        // Generate SAS Uri
                        var sasExpires = DateTime.Now.AddMinutes(30);
                        var sasUri = newBlockBlob.GenerateSasUri(BlobSasPermissions.Read | BlobSasPermissions.Delete, sasExpires);

                        output.Report =  (Name: newBlockBlob.Name, Uri: sasUri.ToString());
                        return output;
                    }

                }
            }
            catch (Exception ex)
            {
                output.Error = ex;
                return output;
            }
        }
    }
}
