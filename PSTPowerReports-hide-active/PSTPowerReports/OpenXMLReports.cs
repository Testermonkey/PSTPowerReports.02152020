using System;
using System.Collections.Generic;
using System.Linq;
using System.IO;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace PSTPowerReports
{
    class OpenXMLReports
    { //OpenXML Excel
        private string outputFileName = string.Empty;
        private const string reportFile = "\\PSTReport.xlsx";

        public OpenXMLReports()
        {
            outputFileName = Directory.GetCurrentDirectory() + reportFile;
        }

        public bool Create(ReportData reportData, List<PowerData> powerLogs)
        {
            bool retStatus = false;
            List<string> cumulativeMemDumpList = new List<string>();
            List<string> cumulativeLiveKernelList = new List<string>();

            try
            {
                if (powerLogs.Count > 0)
                {
                    Util.Trace("Generating report with OpenXML");
                    using (SpreadsheetDocument spreadSheet = SpreadsheetDocument.Create(outputFileName, SpreadsheetDocumentType.Workbook))
                    {
                        SheetData sheetData;
                        if (CreateExcelSheet(spreadSheet, reportData, "PSTSummary", out sheetData))
                        {
                            if (sheetData != null)
                            {
                                // populate data rows
                                foreach (PowerData pData in powerLogs)
                                {
                                    if (pData != null)
                                    {
                                        Write2ExcelSheet(ref sheetData, pData, reportData);
                                        if (pData.MemoryDumpList.Count() > 0)
                                        {
                                            cumulativeMemDumpList.AddRange(pData.MemoryDumpList);
                                        }

                                        if (pData.LiveKernelList.Count() > 0)
                                        {
                                            cumulativeLiveKernelList.AddRange(pData.LiveKernelList);
                                        }
                                    }
                                }
                                WriteReportPath2ExcelSheet(ref sheetData, reportData.CurrPath, "Path: ");
                                WriteDumpList2ExcelSheet(ref sheetData, cumulativeMemDumpList, "Memory dump list:");
                                WriteDumpList2ExcelSheet(ref sheetData, cumulativeLiveKernelList, "Live kernel dump list:");
                            }
                        }
                    }
                    if (File.Exists(outputFileName))
                    {
                        Util.Trace("Report " + outputFileName + " generated");
                        retStatus = true;
                    }
                    else
                    {
                        Util.Trace("OpenXML could not generate report file.\nPlease install Excel/OpenXML and try again!");
                    }
                }
                else
                {
                    Util.Trace(" Nothing to report");
                    retStatus = true;
                }
            }
            catch (Exception ex)
            {
                Util.Trace("Exception in OpenXML::Create -> {0}", ex.Message);
                retStatus = false;
            }
            return retStatus;
        }

        private bool CreateExcelSheet(SpreadsheetDocument spreadSheet, ReportData reportData, string sheetName, out SheetData sheetData)
        {
            bool retStatus = false;
            sheetData = null;
            int cStateColumns = 0;

            cStateColumns += (reportData.C2ActiveAvgPer) ? 1 : 0;
            cStateColumns += (reportData.C3ActiveAvgPer) ? 1 : 0;
            cStateColumns += (reportData.C6ActiveAvgPer) ? 1 : 0;
            cStateColumns += (reportData.C7ActiveAvgPer) ? 1 : 0;
            cStateColumns += (reportData.C8ActiveAvgPer) ? 1 : 0;
            cStateColumns += (reportData.C9ActiveAvgPer) ? 1 : 0;
            cStateColumns += (reportData.C10ActiveAvgPer) ? 1 : 0;
            cStateColumns += (reportData.S0SleepAvgPer) ? 1 : 0;

            try
            {
                WorkbookPart workbookPart = spreadSheet.AddWorkbookPart();
                workbookPart.Workbook = new Workbook();
                WorksheetPart newWorkSheetPart = workbookPart.AddNewPart<WorksheetPart>();
                newWorkSheetPart.Worksheet = new Worksheet();
                Sheets sheets = workbookPart.Workbook.AppendChild(new Sheets());
                Sheet sheet = new Sheet() { Id = workbookPart.GetIdOfPart(newWorkSheetPart), SheetId = 1, Name = "PSTSummary" };
                sheets.Append(sheet);
                workbookPart.Workbook.Save();
                // Adding style
                WorkbookStylesPart stylePart = workbookPart.AddNewPart<WorkbookStylesPart>();
                stylePart.Stylesheet = GenerateStylesheet();
                stylePart.Stylesheet.Save();

                Columns columns = new Columns();
                columns.Append(SetColumnWidth(1, 1, 15.0));
                columns.Append(SetColumnWidth(2, 2, 18.0));
                columns.Append(SetColumnWidth(3, 3, 5.0));
                columns.Append(SetColumnWidth(4, 4, 15.0));
                columns.Append(SetColumnWidth(5, 5, 10.0));
                columns.Append(SetColumnWidth(6, 6, 40.0));
                columns.Append(SetColumnWidth(7, ((uint)cStateColumns + 7), 12.0));
                newWorkSheetPart.Worksheet.Append(columns);

                sheetData = newWorkSheetPart.Worksheet.AppendChild(new SheetData());
                retStatus = CreateHeaderReport(ref sheetData, reportData, cStateColumns);
                if (MergeHeaderCells(out MergeCells mergeHdrCells, cStateColumns))
                {
                    newWorkSheetPart.Worksheet.InsertAfter(mergeHdrCells, newWorkSheetPart.Worksheet.Elements<SheetData>().First());
                }
            }
            catch (Exception e)
            {
                Util.Trace("Exception in OpenXML::CreateExcelSheet -> {0}", e.Message);
                retStatus = false;
            }
            return retStatus;
        }

        private bool CreateHeaderReport(ref SheetData sheetData, ReportData reportData, int cStates)
        {
            // Constructing header
            try
            {
                Row row = new Row();
                string pstVersion = "PST-Wi-Fi v." + reportData.PSTVer;
                if (string.IsNullOrEmpty(reportData.PSTVer))
                {
                    pstVersion = "PST-Wi-Fi v.2.3.1";
                }
                else
                {
                    pstVersion = "PST-Wi-Fi v." + reportData.PSTVer;
                }
                row.Append(
                    ConstructCell(reportData.Header1, CellValues.String, 2),
                    ConstructCell(" ", CellValues.String, 2),
                    ConstructCell(" ", CellValues.String, 2),
                    ConstructCell("", CellValues.String, 2),
                    ConstructCell(pstVersion, CellValues.String, 2),
                    ConstructCell(" ", CellValues.String, 2),
                    ConstructCell("Power Data", CellValues.String, 3));
 
                for (int i = 0; i < cStates; i++)
                {
                    row.Append(ConstructCell(" ", CellValues.String, 2));
                }

                // Insert the header1 row to the Sheet Data
                sheetData.AppendChild(row);
                row = new Row();
                row.Append(
                    ConstructCell(reportData.Header2, CellValues.String, 2),
                    ConstructCell(" ", CellValues.String, 2),
                    ConstructCell(" ", CellValues.String, 2),
                    ConstructCell(" ", CellValues.String, 2),
                    ConstructCell(reportData.CSTimeHeader, CellValues.String, 2),
                    ConstructCell(" ", CellValues.String, 2),
                    ConstructCell("Energy Drain", CellValues.String, 2),
                    ConstructCell("CState Info", CellValues.String, 3));

                for (int i = 0; i < (cStates - 1); i++)
                {
                    row.Append(ConstructCell(" ", CellValues.String, 2));
                }

                // Insert the header2 row to the Sheet Data
                sheetData.AppendChild(row);
                row = new Row();
                row.Append(
                    ConstructCell("Date", CellValues.String, 2),
                    ConstructCell("Device Name", CellValues.String, 2),
                    ConstructCell("OS", CellValues.String, 2),
                    ConstructCell("Power Slider Setting", CellValues.String, 2),
                    ConstructCell("Memory dumped?", CellValues.String, 2),
                    ConstructCell("Bug Check Code", CellValues.String, 2),
                    ConstructCell("Energy Drain Rate", CellValues.String, 2),
                    ConstructCell("Hw Drips %", CellValues.String, 2),
                    ConstructCell("Sw Drips %", CellValues.String, 2));

                if (reportData.C2ActiveAvgPer) row.Append(ConstructCell("% of Time in C2", CellValues.String, 2));
                if (reportData.C3ActiveAvgPer) row.Append(ConstructCell("% of Time in C3", CellValues.String, 2));
                if (reportData.C6ActiveAvgPer) row.Append(ConstructCell("% of Time in C6", CellValues.String, 2));
                if (reportData.C7ActiveAvgPer) row.Append(ConstructCell("% of Time in C7", CellValues.String, 2));
                if (reportData.C8ActiveAvgPer) row.Append(ConstructCell("% of Time in C8", CellValues.String, 2));
                if (reportData.C9ActiveAvgPer) row.Append(ConstructCell("% of Time in C9", CellValues.String, 2));
                if (reportData.C10ActiveAvgPer) row.Append(ConstructCell("% of Time in C10", CellValues.String, 2));
                if (reportData.S0SleepAvgPer) row.Append(ConstructCell("Sleep_S0Per", CellValues.String, 2));

                // Insert the header row to the Sheet Data
                sheetData.AppendChild(row);
            }
            catch(Exception e)
            {
                Util.Trace("Exception in OpenXML::CreateHeaderReport -> {0}", e.Message);
                return false;
            }
            return true;
        }

        private bool MergeHeaderCells(out MergeCells mergeCells, int cStates)
        { //create a MergeCells object to hold each MergeCell
            mergeCells = new MergeCells();

            try
            {
                //append a MergeCell to the mergeCells for each set of merged cells
                mergeCells.Append(new MergeCell() { Reference = new StringValue("A1:D1") });
                mergeCells.Append(new MergeCell() { Reference = new StringValue("E1:F1") });
                mergeCells.Append(new MergeCell() { Reference = 
                        new StringValue(string.Format("G1:{0}1", (char)('H' + (cStates - 1)))) });
                mergeCells.Append(new MergeCell() { Reference = new StringValue("A2:D2") });
                mergeCells.Append(new MergeCell() { Reference = new StringValue("E2:F2") });
                mergeCells.Append(new MergeCell() { Reference = 
                        new StringValue(string.Format("H2:{0}2", (char)('H' + (cStates - 1)))) });
            }
            catch(Exception e)
            {
                Util.Trace("Exception in OpenXML::MergeHeaderCells -> {0}", e.Message);
                return false;
            }
            return true;
        }

        private Cell ConstructCell(string value, CellValues dataType, uint styleIndex = 1)
        {
            return new Cell()
            {
                CellValue = new CellValue(value),
                DataType = new EnumValue<CellValues>(dataType),
                StyleIndex = styleIndex
            };
        }

        private Column SetColumnWidth(uint startColumnIdx, uint endColumnIdx, double columnWidth )
        {
            Column column = new Column
            {
                Min = startColumnIdx,
                Max = endColumnIdx,
                Width = columnWidth,
                CustomWidth = true
            };
            return column;
        }

        private bool Write2ExcelSheet(ref SheetData sheetData, PowerData pData, ReportData reportData)
        {
            try
            {
                string memDumpFlag = string.Empty;
                if ((pData.HasMemoryDump) || (pData.HasLiveKernelDump))
                {
                    memDumpFlag = pData.MemoryDumpFlag + " " + pData.LiveKernelDumpFlag;
                }
                else
                {
                    memDumpFlag = pData.MemoryDumpFlag;
                }
                string bugCheckCodes = string.Empty;
                if (pData.BugCheckCode.Count > 0)
                {
                    for (int i = 0; i < pData.BugCheckCode.Count; i++)
                    {
                        bugCheckCodes += pData.BugCheckCode[i];
                        if (i < (pData.BugCheckCode.Count - 1))
                        {
                            bugCheckCodes += ", ";
                        }
                    }
                }
                if (pData.LiveKernelCode.Count > 0)
                {
                    if (pData.BugCheckCode.Count > 0)
                    {
                        bugCheckCodes += ", ";
                    }
                    for (int i = 0; i < pData.LiveKernelCode.Count; i++)
                    {
                        bugCheckCodes += pData.LiveKernelCode[i];
                        if (i < (pData.LiveKernelCode.Count - 1))
                        {
                            bugCheckCodes += ", ";
                        }
                    }
                }
                Row row = new Row();
                string powerSliderValue = string.Empty;
                if (pData.PowerSourceType != PowerSourceType.None)
                {
                    powerSliderValue = pData.PowerSliderMode.ToString() + " (" + pData.PowerSourceType.ToString() + ")";
                }
                else
                {
                    powerSliderValue = pData.PowerSliderMode.ToString();
                }

                string currDeviceName = pData.DeviceName;
                if (!string.IsNullOrEmpty(pData.RelativePath))
                {
                    currDeviceName += "\n(" + pData.RelativePath + ")";
                }
                row.Append(
                    ConstructCell(pData.Date.ToString(), CellValues.String),
                    ConstructCell(currDeviceName, CellValues.String),
                    ConstructCell(pData.OsType.ToString(), CellValues.String),
                    ConstructCell(powerSliderValue, CellValues.String),
                    ConstructCell(memDumpFlag, CellValues.String),
                    ConstructCell(bugCheckCodes, CellValues.String),
                    ConstructCell(pData.EnergyDrainRate.ToString("0.##"), CellValues.Number),
                    ConstructCell(pData.hwDrips.ToString("0.##"), CellValues.Number),
                    ConstructCell(pData.swDrips.ToString("0.##"), CellValues.Number));

                if (reportData.C2ActiveAvgPer)
                {
                    if (pData.C2ActivePercent != 0.0)
                    {
                        row.Append(ConstructCell(pData.C2ActivePercent.ToString("0.##"), CellValues.Number));
                    }
                    else
                    {
                        row.Append(ConstructCell(" ", CellValues.String));
                    }
                }

                if (reportData.C3ActiveAvgPer)
                {
                    if (pData.C3ActivePercent != 0.0)
                    {
                        row.Append(ConstructCell(pData.C3ActivePercent.ToString("0.##"), CellValues.Number));
                    }
                    else
                    {
                        row.Append(ConstructCell(" ", CellValues.String));
                    }
                }

                if (reportData.C6ActiveAvgPer)
                {
                    if (pData.C6ActivePercent != 0.0)
                    {
                        row.Append(ConstructCell(pData.C6ActivePercent.ToString("0.##"), CellValues.Number));
                    }
                    else
                    {
                        row.Append(ConstructCell(" ", CellValues.String));
                    }
                }

                if (reportData.C7ActiveAvgPer)
                {
                    if (pData.C7ActivePercent != 0.0)
                    {
                        row.Append(ConstructCell(pData.C7ActivePercent.ToString("0.##"), CellValues.Number));
                    }
                    else
                    {
                        row.Append(ConstructCell(" ", CellValues.String));
                    }
                }

                if (reportData.C8ActiveAvgPer)
                {
                    if (pData.C8ActivePercent != 0.0)
                    {
                        row.Append(ConstructCell(pData.C8ActivePercent.ToString("0.##"), CellValues.Number));
                    }
                    else
                    {
                        row.Append(ConstructCell(" ", CellValues.String));
                    }
                }

                if (reportData.C9ActiveAvgPer)
                {
                    if (pData.C9ActivePercent != 0.0)
                    {
                        row.Append(ConstructCell(pData.C9ActivePercent.ToString("0.##"), CellValues.Number));
                    }
                    else
                    {
                        row.Append(ConstructCell(" ", CellValues.String));
                    }
                }

                if (reportData.C10ActiveAvgPer)
                {
                    if (pData.C2ActivePercent != 0.0)
                    {
                        row.Append(ConstructCell(pData.C10ActivePercent.ToString("0.##"), CellValues.Number));
                    }
                    else
                    {
                        row.Append(ConstructCell(" ", CellValues.String));
                    }
                }

                if (reportData.S0SleepAvgPer)
                {
                    if (pData.SleepS0Percent != 0.0)
                    {
                        row.Append(ConstructCell(pData.SleepS0Percent.ToString("0.##"), CellValues.Number));
                    }
                    else
                    {
                        row.Append(ConstructCell(" ", CellValues.String));
                    }
                }
                if (reportData.SwDrips)
                {
                    if (pData.swDrips != 0.0)
                    {
                        row.Append(ConstructCell(pData.swDrips.ToString("0.##"), CellValues.Number));
                    }
                    else
                    {
                        row.Append(ConstructCell(" ", CellValues.String));
                    }
                }
                if (reportData.HwDrips)
                {
                    if (pData.hwDrips != 0.0)
                    {
                        row.Append(ConstructCell(pData.hwDrips.ToString("0.##"), CellValues.Number));
                    }
                    else
                    {
                        row.Append(ConstructCell(" ", CellValues.String));
                    }
                }

                sheetData.AppendChild(row);
            }
            catch (Exception e)
            {
                Util.Trace("Exception in OpenXML->Write2ExcelSheet = {0}", e.Message);
                return false;
            }

            return true;
        }

        private bool WriteDumpList2ExcelSheet(ref SheetData sheetData, List<string> dmpList, string header)
        {
            try
            {
                Row row = new Row();
                row.Append(ConstructCell("", CellValues.String, 4));
                sheetData.AppendChild(row);
                row = new Row();
                row.Append(ConstructCell("", CellValues.String, 4));
                sheetData.AppendChild(row);
                row = new Row();
                row.Append(ConstructCell(header, CellValues.String, 5));
                sheetData.AppendChild(row);

                int idx = 1;
                foreach (string list in dmpList)
                {
                    row = new Row
                    {
                        StyleIndex = Convert.ToUInt32(0)
                    };
                    row.Append(ConstructCell(string.Format("    {0})  {1}", idx++, list), CellValues.String, 4));
                    sheetData.AppendChild(row);
                }
            }
            catch (Exception e)
            {
                Util.Trace("Exception in OpenXML->WriteDumpList2ExcelSheet = {0}", e.Message);
                return false;
            }

            return true;
        }

        private bool WriteReportPath2ExcelSheet(ref SheetData sheetData, string reportPath, string header)
        {
            try
            {
                Row row = new Row();
                row.Append(ConstructCell("", CellValues.String, 4));
                sheetData.AppendChild(row);
                row = new Row();
                row.Append(ConstructCell(header + reportPath, CellValues.String, 5));
                sheetData.AppendChild(row);

            }
            catch (Exception e)
            {
                Util.Trace("Exception in OpenXML->WriteDumpList2ExcelSheet = {0}", e.Message);
                return false;
            }

            return true;
        }

        private Stylesheet GenerateStylesheet()
        {
            Stylesheet styleSheet = null;

            Fonts fonts = new Fonts(
                new Font( // Index 0 - default
                    new FontSize() { Val = 10 }),
                new Font( // Index 1 - header
                    new FontSize() { Val = 10 }, new Bold()/*, new Color() { Rgb = "FFFFFF" }*/));

            Fills fills = new Fills(
                    new Fill(new PatternFill() { PatternType = PatternValues.None }), // Index 0 - default
                    new Fill(new PatternFill() { PatternType = PatternValues.Gray125 }), // Index 1 - default
                    new Fill(new PatternFill(new ForegroundColor { Rgb = new HexBinaryValue() { Value = "66666666" } })
                    { PatternType = PatternValues.Solid }) // Index 2 - header
                    );

            Borders borders = new Borders(
                    new Border(), // index 0 default
                    new Border( // index 1 black border
                        new LeftBorder(new Color() { Auto = true }) { Style = BorderStyleValues.Thin },
                        new RightBorder(new Color() { Auto = true }) { Style = BorderStyleValues.Thin },
                        new TopBorder(new Color() { Auto = true }) { Style = BorderStyleValues.Thin },
                        new BottomBorder(new Color() { Auto = true }) { Style = BorderStyleValues.Thin },
                        new DiagonalBorder())
                    );

            CellFormats cellFormats = new CellFormats(
                    new CellFormat(), // default
                    new CellFormat(new Alignment() { WrapText = true })
                    { FontId = 0, FillId = 0, BorderId = 1, ApplyBorder = true, ApplyAlignment = true }, // body
                    new CellFormat(new Alignment() { WrapText = true })
                    { FontId = 1, FillId = 0, BorderId = 1, ApplyBorder = true, ApplyAlignment = true }, // header
                    new CellFormat(new Alignment() { Horizontal = HorizontalAlignmentValues.Center, WrapText = true })
                    { FontId = 1, FillId = 0, BorderId = 1, ApplyBorder = true },
                    new CellFormat(new Alignment() { WrapText = false })
                    { FontId = 0, FillId = 0, BorderId = 0, ApplyBorder = false },
                    new CellFormat(new Alignment() { WrapText = false })
                    { FontId = 1, FillId = 0, BorderId = 0, ApplyBorder = false });

            styleSheet = new Stylesheet(fonts, fills, borders, cellFormats);

            return styleSheet;
        }
    }
}