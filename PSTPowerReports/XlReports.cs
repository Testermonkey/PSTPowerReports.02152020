using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;

namespace PSTPowerReports
{
    public enum CellTextAlign
    {
        left = 0,
        right = 1,
        center = 2,
    }

    class XlReports
    { // MS Office Excel
        private bool ExcelSheetNotEmpty = false;
        private string outputFileName = string.Empty;
        private const string reportFile = "\\PSTReport.xlsx";

        public XlReports()
        {
            ExcelSheetNotEmpty = false;
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
                    Util.Trace("Generating report with MS Office Excel");
                    Excel.Application xlApp = new Excel.Application();
                    if (xlApp != null)
                    {
                        if (File.Exists(outputFileName))
                        {
                            File.Delete(outputFileName);
                        }
                        Excel.Workbook xlWorkbook = xlApp.Workbooks.Add();
                        if (CreateExcelSheet(xlWorkbook, reportData, "PSTSummary"))
                        {
                            Excel.Worksheet xlWorksheet = (Excel.Worksheet)xlWorkbook.Sheets["PSTSummary"];
                            int rowNum = GetLastRowIndex(xlWorksheet);

                            foreach (PowerData pData in powerLogs)
                            {
                                if (pData != null)
                                {
                                    Write2ExcelSheet(xlWorksheet, ref rowNum, pData, reportData);
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
                            int colCount = 14;
                            colCount += (reportData.C2ActiveAvgPer) ? 1 : 0;
                            colCount += (reportData.C3ActiveAvgPer) ? 1 : 0;
                            colCount += (reportData.C6ActiveAvgPer) ? 1 : 0;
                            colCount += (reportData.C7ActiveAvgPer) ? 1 : 0;
                            colCount += (reportData.C8ActiveAvgPer) ? 1 : 0;
                            colCount += (reportData.C9ActiveAvgPer) ? 1 : 0;
                            colCount += (reportData.C10ActiveAvgPer) ? 1 : 0;
                            AddBorders(xlWorksheet, rowNum, colCount);

                            WriteReportPath2ExcelSheet(xlWorksheet, ref rowNum, reportData.CurrPath, "Path: ");
                            WriteDumpList2ExcelSheet(xlWorksheet, ref rowNum, cumulativeMemDumpList, "Memory dump list:");
                            WriteDumpList2ExcelSheet(xlWorksheet, ref rowNum, cumulativeLiveKernelList, "Live kernel dump list:");

                            //release com objects to fully kill excel process from running in the background
                            Marshal.ReleaseComObject(xlWorksheet);
                            ExcelSheetNotEmpty = true;

                            xlWorkbook.SaveAs(outputFileName, Microsoft.Office.Interop.Excel.XlFileFormat.xlWorkbookDefault,
                                                Type.Missing, Type.Missing, false, false,
                                                Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange,
                                                Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                        }

                        //cleanup
                        GC.Collect();
                        GC.WaitForPendingFinalizers();

                        //close and release com objects
                        xlWorkbook.Close();
                        Marshal.ReleaseComObject(xlWorkbook);

                        //quit and release
                        xlApp.Quit();
                        Marshal.ReleaseComObject(xlApp);

                        if (File.Exists(outputFileName))
                        {
                            Util.Trace("Report " + outputFileName + " generated");
                            retStatus = true;
                        }
                        else
                        {
                            Util.Trace("Excel application is required for report generation.\nPlease install Excel and try again!");
                        }
                    }
                    else
                    {
                        Util.Trace("EXCEL could not be started. Check your office installation");
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
                Util.Trace("Exception in XlReports::Create -> {0}",ex.Message);
                retStatus = false;
            }
            return retStatus;
        }

        private bool CreateExcelSheet(Excel.Workbook xlWorkbook, ReportData reportData, string sheetName)
        {
            try
            {
                Excel.Worksheet xlWorksheet = (Excel.Worksheet)xlWorkbook.Sheets["sheet1"];
                if (xlWorksheet == null)
                {
                    return false;
                }
                xlWorksheet.Name = sheetName;
                FormatColumn(xlWorksheet, 1, 15);
                FormatColumn(xlWorksheet, 2, 18);
                FormatColumn(xlWorksheet, 3, 5);
                FormatColumn(xlWorksheet, 4, 15);
                FormatColumn(xlWorksheet, 5, 10);
                FormatColumn(xlWorksheet, 6, 40);
                FormatColumn(xlWorksheet, 7, 12);
                FormatColumn(xlWorksheet, 8, 12);
                FormatColumn(xlWorksheet, 9, 12);
                FormatColumn(xlWorksheet, 10, 12);
                FormatColumn(xlWorksheet, 11, 12);
                FormatColumn(xlWorksheet, 12, 12);
                FormatColumn(xlWorksheet, 13, 12);
                int colCount = 13;
                FormatColumn(xlWorksheet, (colCount += ((reportData.C2ActiveAvgPer) ? 1 : 0)), 12);
                FormatColumn(xlWorksheet, (colCount += ((reportData.C3ActiveAvgPer) ? 1 : 0)), 12);
                FormatColumn(xlWorksheet, (colCount += ((reportData.C6ActiveAvgPer) ? 1 : 0)), 12);
                FormatColumn(xlWorksheet, (colCount += ((reportData.C7ActiveAvgPer) ? 1 : 0)), 12);
                FormatColumn(xlWorksheet, (colCount += ((reportData.C8ActiveAvgPer) ? 1 : 0)), 12);
                FormatColumn(xlWorksheet, (colCount += ((reportData.C9ActiveAvgPer) ? 1 : 0)), 12);
                FormatColumn(xlWorksheet, (colCount += ((reportData.C10ActiveAvgPer) ? 1 : 0)), 12);
                FormatColumn(xlWorksheet, (colCount += ((reportData.S0SleepAvgPer) ? 1 : 0)), 12);
                CreateHeaderReport(xlWorksheet, reportData);
            }
            catch (Exception e)
            {
                Util.Trace("Exception in XlReports::CreateExcelSheet -> {0}", e.Message);
                return false;
            }
            return true;
        }

        private bool Write2ExcelSheet(Excel.Worksheet xlWorksheet, ref int rowNum, PowerData pData, ReportData reportData)
        {
            try
            {
                if (rowNum >= 1)
                {
                    rowNum++;
                    FillCellValue(xlWorksheet, rowNum, 1, pData.Date.ToString());
                    string currDeviceName = pData.DeviceName;
                    if (!string.IsNullOrEmpty(pData.RelativePath))
                    {
                        currDeviceName += "\n(" + pData.RelativePath + ")";
                    }
                    FillCellValue(xlWorksheet, rowNum, 2, currDeviceName);
                    FillCellValue(xlWorksheet, rowNum, 3, pData.OsType.ToString());
                    if (pData.PowerSourceType != PowerSourceType.None)
                    {
                        FillCellValue(xlWorksheet, rowNum, 4, pData.PowerSliderMode.ToString() + " (" + pData.PowerSourceType.ToString() + ")");
                    }
                    else
                    {
                        FillCellValue(xlWorksheet, rowNum, 4, pData.PowerSliderMode.ToString());
                    }
                    string memDumpFlag = string.Empty;
                    if ((pData.HasMemoryDump) || (pData.HasLiveKernelDump))
                    {
                        memDumpFlag = pData.MemoryDumpFlag + " " + pData.LiveKernelDumpFlag;
                    }
                    else
                    {
                        memDumpFlag = pData.MemoryDumpFlag;
                    }
                    FillCellValue(xlWorksheet, rowNum, 5, memDumpFlag);
                    string bugCheckCodes = string.Empty;
                    if (pData.BugCheckCode.Count > 0)
                    {
                        for(int i = 0; i < pData.BugCheckCode.Count; i++)
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
                    FillCellValue(xlWorksheet, rowNum, 6, bugCheckCodes);
                    FillCellValue(xlWorksheet, rowNum, 7, pData.EnergyDrainRate.ToString("0.##"), false, CellTextAlign.right);
                    FillCellValue(xlWorksheet, rowNum, 8, pData.hwDrips.ToString("0.##"), false, CellTextAlign.right);
                    FillCellValue(xlWorksheet, rowNum, 9, pData.swDrips.ToString("0.##"), false, CellTextAlign.right);
                    FillCellValue(xlWorksheet, rowNum, 10, pData.TotalBatt.ToString("0.##"), false, CellTextAlign.right);
                    FillCellValue(xlWorksheet, rowNum, 11, pData.Batt1.ToString("0.##"), false, CellTextAlign.right);
                    FillCellValue(xlWorksheet, rowNum, 12, pData.Batt2.ToString("0.##"), false, CellTextAlign.right);
                    FillCellValue(xlWorksheet, rowNum, 13, pData.ActiveEnergyDrainRate.ToString("0.##"), false, CellTextAlign.right);

                    int colCount = 13;
                    if (reportData.C2ActiveAvgPer)
                    {
                        ++colCount;
                        if (pData.C2ActivePercent != 0.0)
                        {
                            FillCellValue(xlWorksheet, rowNum, colCount, pData.C2ActivePercent.ToString("0.##"), false, CellTextAlign.right);
                        }
                    }
                    if (reportData.C3ActiveAvgPer)
                    {
                        ++colCount;
                        if (pData.C3ActivePercent != 0.0)
                        {
                            FillCellValue(xlWorksheet, rowNum, colCount, pData.C3ActivePercent.ToString("0.##"), false, CellTextAlign.right);
                        }
                    }
                    if (reportData.C6ActiveAvgPer)
                    {
                        ++colCount;
                        if (pData.C6ActivePercent != 0.0)
                        {
                            FillCellValue(xlWorksheet, rowNum, colCount, pData.C6ActivePercent.ToString("0.##"), false, CellTextAlign.right);
                        }
                    }
                    if (reportData.C7ActiveAvgPer)
                    {
                        ++colCount;
                        if (pData.C7ActivePercent != 0.0)
                        {
                            FillCellValue(xlWorksheet, rowNum, colCount, pData.C7ActivePercent.ToString("0.##"), false, CellTextAlign.right);
                        }
                    }
                    if (reportData.C8ActiveAvgPer)
                    {
                        ++colCount;
                        if (pData.C8ActivePercent != 0.0)
                        {
                            FillCellValue(xlWorksheet, rowNum, colCount, pData.C8ActivePercent.ToString("0.##"), false, CellTextAlign.right);
                        }
                    }
                    if (reportData.C9ActiveAvgPer)
                    {
                        ++colCount;
                        if (pData.C9ActivePercent != 0.0)
                        {
                            FillCellValue(xlWorksheet, rowNum, colCount, pData.C9ActivePercent.ToString("0.##"), false, CellTextAlign.right);
                        }
                    }
                    if (reportData.C10ActiveAvgPer)
                    {
                        ++colCount;
                        if (pData.C10ActivePercent != 0.0)
                        {
                            FillCellValue(xlWorksheet, rowNum, colCount, pData.C10ActivePercent.ToString("0.##"), false, CellTextAlign.right);
                        }
                    }
                    if (reportData.S0SleepAvgPer)
                    {
                        ++colCount;
                        if (pData.SleepS0Percent != 0.0)
                        {
                            FillCellValue(xlWorksheet, rowNum, colCount, pData.SleepS0Percent.ToString("0.##"), false, CellTextAlign.right);
                        }
                    }
                    if (reportData.BatteryData)
                    {
                        ++colCount;
                        if (pData.TotalBatt != 0.0)
                        {
                            FillCellValue(xlWorksheet, rowNum, colCount, pData.TotalBatt.ToString("0.##"), false, CellTextAlign.right);
                        }
                        ++colCount;
                        if (pData.Batt1 != 0.0)
                        {
                            FillCellValue(xlWorksheet, rowNum, colCount, pData.Batt1.ToString("0.##"), false, CellTextAlign.right);
                        }
                        ++colCount;
                        if (pData.Batt2 != 0.0)
                        {
                            FillCellValue(xlWorksheet, rowNum, colCount, pData.Batt2.ToString("0.##"), false, CellTextAlign.right);
                        }
                    }
                    if (reportData.ActiveEnergyDrain)
                    {
                        ++colCount;
                        if (pData.ActiveEnergyDrainRate != 0.0)
                        {
                            FillCellValue(xlWorksheet, rowNum, colCount, pData.ActiveEnergyDrainRate.ToString("0.##"), false, CellTextAlign.right);
                        }
                    }
                }
            }
            catch (Exception e)
            {
                Util.Trace("Exception in XlReports->Write2ExcelSheet = {0}", e.Message);
                return false;
            }

            return true;
        }

        private bool WriteDumpList2ExcelSheet(Excel.Worksheet xlWorksheet, ref int rowNum, List<string> dumpList, string header)
        {
            try
            {
                if ((rowNum >= 1) && (dumpList.Count() > 0))
                {
                    rowNum += 3;
                    FillRowValue(xlWorksheet, rowNum, 1, header, true);
                    int idx = 1;
                    foreach (string list in dumpList)
                    {
                        rowNum++;
                        FillRowValue(xlWorksheet, rowNum, 1, string.Format("    {0})  {1}",idx++, list));
                    }
                }
            }
            catch (Exception e)
            {
                Util.Trace("Exception in XlReports->WriteDumpList2ExcelSheet = {0}", e.Message);
                return false;
            }

            return true;
        }

        private bool WriteReportPath2ExcelSheet(Excel.Worksheet xlWorksheet, ref int rowNum, string reportPath, string header)
        {
            try
            {
                if ((rowNum >= 1) && (!string.IsNullOrEmpty(reportPath)) && (!string.IsNullOrEmpty(header)))
                {
                    rowNum += 2;
                    FillRowValue(xlWorksheet, rowNum, 1, (header + reportPath), true);
                }
            }
            catch (Exception e)
            {
                Util.Trace("Exception in XlReports->WriteReportPath2ExcelSheet = {0}", e.Message);
                return false;
            }

            return true;
        }

        private void FormatColumn(Excel.Worksheet wkSheet, int colNum, int width)
        {
            wkSheet.Columns[colNum].ColumnWidth = width;
            wkSheet.Columns[colNum].VerticalAlignment = true;
            wkSheet.Columns[colNum].WrapText = true;
        }

        private void CreateHeaderReport(Excel.Worksheet wkSheet, ReportData reportData)
        {
            int rowCount = GetLastRowIndex(wkSheet);
            rowCount += 2;

            int cStateColumns = 0;
            cStateColumns += (reportData.C2ActiveAvgPer) ? 1 : 0;
            cStateColumns += (reportData.C3ActiveAvgPer) ? 1 : 0;
            cStateColumns += (reportData.C6ActiveAvgPer) ? 1 : 0;
            cStateColumns += (reportData.C7ActiveAvgPer) ? 1 : 0;
            cStateColumns += (reportData.C8ActiveAvgPer) ? 1 : 0;
            cStateColumns += (reportData.C9ActiveAvgPer) ? 1 : 0;
            cStateColumns += (reportData.C10ActiveAvgPer) ? 1 : 0;
            cStateColumns += (reportData.S0SleepAvgPer) ? 1 : 0;
            // Header first two rows formatting
            wkSheet.Range[wkSheet.Cells[1, 1], wkSheet.Cells[1, 6]].Merge();
            wkSheet.Range[wkSheet.Cells[1, 7], wkSheet.Cells[1, 13]].Merge();
            wkSheet.Range[wkSheet.Cells[1, 14], wkSheet.Cells[1, (14 + cStateColumns)]].Merge();
            wkSheet.Range[wkSheet.Cells[2, 1], wkSheet.Cells[2, 4]].Merge();
            wkSheet.Range[wkSheet.Cells[2, 5], wkSheet.Cells[2, 6]].Merge();
            wkSheet.Range[wkSheet.Cells[2, 7], wkSheet.Cells[2, 7]].Merge();
            wkSheet.Range[wkSheet.Cells[2, 8], wkSheet.Cells[2, 9]].Merge();
            wkSheet.Range[wkSheet.Cells[2, 10], wkSheet.Cells[2, 12]].Merge();
            wkSheet.Range[wkSheet.Cells[2, 13], wkSheet.Cells[2, 13]].Merge();
            wkSheet.Range[wkSheet.Cells[2, 14], wkSheet.Cells[2, (14 + cStateColumns)]].Merge();
            FillCellValue(wkSheet, 1, 1, reportData.Header1, true);
            string pstVersion = "PST-Wi-Fi v." + reportData.PSTVer;
            if (string.IsNullOrEmpty(reportData.PSTVer))
            {
                pstVersion = "PST-Wi-Fi v.2.3.1";
            }
            else
            {
                pstVersion = "PST-Wi-Fi v." + reportData.PSTVer;
            }
            FillCellValue(wkSheet, 1, 5, pstVersion, true);
            FillCellValue(wkSheet, 2, 1, reportData.Header2, true);
            FillCellValue(wkSheet, 2, 5, reportData.CSTimeHeader, true);
            FillCellValue(wkSheet, 1, 7, "Power Data", true, CellTextAlign.center);
            FillCellValue(wkSheet, 2, 7, "Energy Drain", true, CellTextAlign.center);
            FillCellValue(wkSheet, 2, 8, "Drips", true, CellTextAlign.center);
            FillCellValue(wkSheet, 2, 10, "Battery Charge Info", true, CellTextAlign.center);
            FillCellValue(wkSheet, 2, 13, "Active Power", true, CellTextAlign.center);
            FillCellValue(wkSheet, 2, 14, "CState Info", true, CellTextAlign.center);

            if (rowCount >= 2)
            {
                rowCount++;
                wkSheet.Cells[rowCount, 1].EntireRow.Font.Bold = true;

                FillCellValue(wkSheet, rowCount, 1, "Date");
                FillCellValue(wkSheet, rowCount, 2, "Device Name");
                FillCellValue(wkSheet, rowCount, 3, "OS");
                FillCellValue(wkSheet, rowCount, 4, "Power Slider Setting");
                FillCellValue(wkSheet, rowCount, 5, "Memory dumped?");
                FillCellValue(wkSheet, rowCount, 6, "Bug Check Code");
                FillCellValue(wkSheet, rowCount, 7, "Energy Drain Rate");
                FillCellValue(wkSheet, rowCount, 8, "Hw Drip %");
                FillCellValue(wkSheet, rowCount, 9, "Sw Drip %");
                FillCellValue(wkSheet, rowCount, 10, "Total Battery %");
                FillCellValue(wkSheet, rowCount, 11, "Bat01 %");
                FillCellValue(wkSheet, rowCount, 12, "Bat02 %");
                FillCellValue(wkSheet, rowCount, 13, "Drain Rate mWh/h");
                int colCount = 13;
                if (reportData.C2ActiveAvgPer) FillCellValue(wkSheet, rowCount, (colCount += ((reportData.C2ActiveAvgPer) ? 1 : 0)), "% of Time in C2");
                if (reportData.C3ActiveAvgPer) FillCellValue(wkSheet, rowCount, (colCount += ((reportData.C3ActiveAvgPer) ? 1 : 0)), "% of Time in C3");
                if (reportData.C6ActiveAvgPer) FillCellValue(wkSheet, rowCount, (colCount += ((reportData.C6ActiveAvgPer) ? 1 : 0)), "% of Time in C6");
                if (reportData.C7ActiveAvgPer) FillCellValue(wkSheet, rowCount, (colCount += ((reportData.C7ActiveAvgPer) ? 1 : 0)), "% of Time in C7");
                if (reportData.C8ActiveAvgPer) FillCellValue(wkSheet, rowCount, (colCount += ((reportData.C8ActiveAvgPer) ? 1 : 0)), "% of Time in C8");
                if (reportData.C9ActiveAvgPer) FillCellValue(wkSheet, rowCount, (colCount += ((reportData.C9ActiveAvgPer) ? 1 : 0)), "% of Time in C9");
                if (reportData.C10ActiveAvgPer) FillCellValue(wkSheet, rowCount, (colCount += ((reportData.C10ActiveAvgPer) ? 1 : 0)), "% of Time in C10");
                if (reportData.S0SleepAvgPer) FillCellValue(wkSheet, rowCount, (colCount += ((reportData.S0SleepAvgPer) ? 1 : 0)), "Sleep_S0Per");
                ExcelSheetNotEmpty = true;
            }
        }

        private int GetLastRowIndex(Excel.Worksheet wkSheet)
        {
            int lastRow = -1;

            try
            {
                if (ExcelSheetNotEmpty)
                {
                    lastRow = wkSheet.Cells.Find(What: (object)"*", LookIn: Excel.XlFindLookIn.xlValues,
                                LookAt: Excel.XlLookAt.xlWhole, SearchOrder: Excel.XlSearchOrder.xlByRows,
                                SearchDirection: Excel.XlSearchDirection.xlPrevious, MatchCase: (object)false).Row;
                }
                else
                {
                    lastRow = 0;
                }
            }
            catch (Exception e)
            {
                Util.Trace("Exception in XlReports::GetLastRow = {0}", e.Message);
            }

            return lastRow;
        }

        private void FillCellValue(Excel.Worksheet wkSheet, int row, int column, string value, bool isBold = false, CellTextAlign txtAlign = CellTextAlign.left)
        {
            wkSheet.Cells[row, column] = value;
            if (txtAlign == CellTextAlign.center)
            {
                wkSheet.Cells[row, column].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
            }
            else if (txtAlign == CellTextAlign.right)
            {
                wkSheet.Cells[row, column].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignRight;
            }
            else if (txtAlign == CellTextAlign.left)
            {
                wkSheet.Cells[row, column].HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignLeft;
            }

            if (isBold)
            {
                wkSheet.Cells[row, column].Font.Bold = true;
            }
        }

        private void FillRowValue(Excel.Worksheet wkSheet, int row, int column, string value, bool isBold = false, CellTextAlign txtAlign = CellTextAlign.left)
        {
            wkSheet.Columns[1].WrapText = false;
            FillCellValue(wkSheet, row, column, value, isBold, txtAlign);
        }

        private void AddBorders(Excel.Worksheet wkSheet, int row, int column)
        {
            Excel.Range range = wkSheet.Range["a1", string.Format("{0}{1}", (char)('a' + (column - 1)), row)];
            Excel.Borders border = range.Borders;
            border.LineStyle = Excel.XlLineStyle.xlContinuous;
            border.Weight = 2d;
        }
    }
}