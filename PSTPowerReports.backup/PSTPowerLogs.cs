using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace PSTPowerReports
{
    class PSTPowerLogs
    {
        private List<PowerData> powerLogs;
        private ReportData reportData;

        public PSTPowerLogs()
        {
            powerLogs = new List<PowerData>();
            reportData = new ReportData();
        }

        public int Initialize(string[] args)
        {
            int retStatus = 0; 
            reportData.Header1 = string.Empty;
            reportData.Header2 = string.Empty;
            string reportinfo = null;
            
            try
            {
                Util.Initialize();
                if (args.Count() <= 0)
                {
                    DisplayHelp();
                    return 0;
                }
                else
                {
                    foreach (string cmd in args)
                    {
                        if (string.Compare(cmd.ToLower(), 0, "/?", 0, 2) == 0)
                        {
                            DisplayHelp();
                            return 0;
                        }
                        else if (string.Compare(cmd.ToLower(), 0, "/h1=", 0, 4) == 0)
                        {
                            string[] cmds = cmd.Split('=');
                            if (cmds.Count() > 1)
                            {
                                reportData.Header1 = cmds[1];
                                reportinfo += cmds[1];
                            }
                        }
                        else if (string.Compare(cmd.ToLower(), 0, "/h2=", 0, 4) == 0)
                        {
                            string[] cmds = cmd.Split('=');
                            if (cmds.Count() > 1)
                            {
                                reportData.Header2 = cmds[1];
                                reportinfo += cmds[1];
                            }
                        }
                        else if (string.Compare(cmd.ToLower(), 0, "/pstver=", 0, 8) == 0)
                        {
                            string[] cmds = cmd.Split('=');
                            if (cmds.Count() > 1)
                            {
                                reportData.PSTVer = cmds[1];
                                reportinfo += cmds[1];
                            }
                        }
                        else if (string.Compare(cmd.ToLower(), 0, "/dir=", 0, 5) == 0)
                        {
                            string[] cmds = cmd.Split('=');
                            if (cmds.Count() > 1)
                            {
                                string path4PST = cmds[1].TrimEnd(' ').TrimEnd('\\');
                                var drive = Path.GetPathRoot(path4PST);
                                if (Environment.GetLogicalDrives().Contains(drive, StringComparer.InvariantCultureIgnoreCase))
                                {
                                    reportData.CurrPath = path4PST;
                                    retStatus = 1;
                                }
                                else if (path4PST.StartsWith(@"\") && !path4PST.StartsWith(@"\\"))
                                {
                                    string wd = Directory.GetCurrentDirectory() + path4PST;
                                    if (Directory.Exists(wd))
                                    {
                                        reportData.CurrPath = wd;
                                        reportinfo += wd;
                                        retStatus = 1;
                                    }
                                    else
                                    {
                                        retStatus = -1;
                                    }
                                }
                                else if (!path4PST.StartsWith(@"\\"))
                                {
                                    string wd = Directory.GetCurrentDirectory() + @"\" + path4PST.TrimStart();
                                    if (Directory.Exists(wd))
                                    {
                                        reportData.CurrPath = wd;
                                        reportinfo += wd;
                                        retStatus = 1;
                                    }
                                    else
                                    {
                                        retStatus = -1;
                                    }
                                }
                                else if (path4PST.StartsWith(@"\\"))
                                {
                                    if (Directory.Exists(Directory.GetCurrentDirectory() + path4PST))
                                    { // to avoid reporting incorrect path
                                        reportData.CurrPath = Directory.GetCurrentDirectory() 
                                                                + @"\" + path4PST.TrimStart('\\');
                                        reportinfo += reportData.CurrPath;
                                    }
                                    else
                                    { // to avoid reporting incorrect path
                                        reportData.CurrPath = @"\\" + path4PST.TrimStart('\\');
                                        reportinfo += reportData.CurrPath;
                                    }
                                    retStatus = 1;
                                }
                                else
                                {
                                    retStatus = -1;
                                }
                                reportData.CurrPath = Path.GetFullPath(reportData.CurrPath);
                                if (string.IsNullOrEmpty(reportData.CurrPath))
                                {
                                    retStatus = -1;
                                }
                            }
                        }
                        else
                        {
                            Util.Trace("Error: Incorrect command parameters");
                            retStatus = -1;
                            break;
                        }
                    }
                }

                if (retStatus <= 0)
                {
                    if (retStatus == 0)
                        Util.Trace("Error: PST test results folder wasn't specified");
                    DisplayHelp();
                }
            }
            catch(Exception e)
            {
                Util.Trace("Exception in PSTPowerLogs::Initialize -> {0}", e.Message);
                retStatus = -2;
            }
            Console.WriteLine("Processing reports with params: {0}", reportinfo);
            return retStatus;
        }

        public bool Process(out bool exceptionGenerated)
        {
            bool retStatus = false;
            exceptionGenerated = false;

            try
            {
                if (Directory.Exists(reportData.CurrPath))
                {
                    DirectoryInfo pstTestResultsDir = new DirectoryInfo(reportData.CurrPath);
                    DirectoryInfo[] dirs = pstTestResultsDir.GetDirectories();
                    List<string> pstDirs = new List<string>();

                    GetPSTFolders(dirs, ref pstDirs);
                    reportData.Process(pstDirs);
                    foreach (string currDir in pstDirs)
                    {
                        if (Directory.Exists(currDir))
                        {
                            if (!GetPowerData(currDir, out PowerData powerData))
                            {
                                Util.Trace("Warning: Could not get Power data for {0}", currDir);
                            }
                            else
                            {
                                powerLogs.Add(powerData);
                                retStatus = true;
                            }
                        }
                    }

                    if (!retStatus)
                    { // try finding PST test results in current directory 
                        if (CheckForPSTResults(reportData.CurrPath))
                        {
                            DirectoryInfo currPSTTestResultsDir = new DirectoryInfo(reportData.CurrPath);
                            if (!GetPowerData(currPSTTestResultsDir.FullName, out PowerData powerData))
                            {
                                Util.Trace("Warning: Could not get Power data for {0}", currPSTTestResultsDir.Name);
                            }
                            else
                            {
                                powerLogs.Add(powerData);
                                retStatus = true;
                            }
                        }
                    }
                }
            }
            catch(Exception e)
            {
                Util.Trace("Exception in PSTPowerLogs::Process -> {0}", e.Message);
                retStatus = false;
                exceptionGenerated = true;
            }

            return retStatus;
        }

        public bool GenerateReport(out bool exceptionGenerated)
        {
            exceptionGenerated = false;

            try
            {
                Type officeType = Type.GetTypeFromProgID("Excel.Application");
                if (officeType != null)
                { // MS Office Excel
                    return (new XlReports().Create(reportData, powerLogs));
                }
                else
                { // using OpenXML Excel
                    return (new OpenXMLReports().Create(reportData, powerLogs));
                }
            }
            catch(Exception e)
            {
                Util.Trace("Exception in PSTPowerLogs::GenerateReport -> {0}", e.Message);
                exceptionGenerated = true;
                return false;
            }
        }

        private bool GetPowerData(string pstDeviceFolder, out PowerData powerData)
        {
            powerData = new PowerData();
            if (string.IsNullOrEmpty(pstDeviceFolder))
            {
                return false;
            }

            if (PSTDataProcessing.GetDeviceDateAndName(pstDeviceFolder.Split(Path.DirectorySeparatorChar).Last(), out DateTime tDate, out string deviceName))
            {
                powerData.Date = tDate;
                powerData.DeviceName = deviceName;
                if (pstDeviceFolder.Length > reportData.CurrPath.Length)
                {
                    string remainPath = pstDeviceFolder.Substring(reportData.CurrPath.Length);
                    int relativePathLength = remainPath.Length - (pstDeviceFolder.Split(Path.DirectorySeparatorChar).Last().Length + 1);
                    if (relativePathLength > 0)
                    {
                        powerData.RelativePath = remainPath.Substring(0, relativePathLength);
                    }
                }
            }

            if (PSTDataProcessing.GetOSType(pstDeviceFolder, out string osType))
            {
                powerData.OsType = osType;
            }

            if (PSTDataProcessing.GetDrips(pstDeviceFolder, out double HwDrips, out double SwDrips, out double aedr))
            {
                powerData.hwDrips = HwDrips;
                powerData.swDrips = SwDrips;
                powerData.ActiveEnergyDrainRate = aedr;
            }

            if (PSTDataProcessing.GetPowerSliderTypeFromDevice(pstDeviceFolder, out bool ACFlag, out PowerSliderType pSliderFromDevice, out double bateryTotal, out double bat1level, out double bat2level))
            {
                powerData.TotalBatt = bateryTotal;
                powerData.Batt1 = bat1level;
                powerData.Batt2 = bat2level;
                powerData.PowerSliderMode = pSliderFromDevice;
                powerData.PowerSourceType = ACFlag ? PowerSourceType.AC : PowerSourceType.Battery;
            }
            else if (PSTDataProcessing.GetPowerSliderType(pstDeviceFolder, out PowerSliderType pSlider))
            {
                powerData.PowerSliderMode = pSlider;
            }

            if (PSTDataProcessing.CheckMemoryDump(pstDeviceFolder, out bool hasMemDmp, out List<string> currCrashDetails, out List<string> memDumpList))
            {
                powerData.HasMemoryDump = hasMemDmp;
                if (hasMemDmp)
                {
                    powerData.MemoryDumpFlag = "Yes";
                    powerData.BugCheckCode = currCrashDetails;
                }
                else
                {
                    powerData.MemoryDumpFlag = "No";
                }

                if (memDumpList.Count() > 0)
                {
                    powerData.MemoryDumpList = memDumpList;
                }
            }

            if (PSTDataProcessing.CheckLiveKernelDump(pstDeviceFolder, out bool hasLiveKernelDmp, out List<string> currLiveKernelDetails, out List<string> liveKernelReportList))
            {
                powerData.HasLiveKernelDump = hasLiveKernelDmp;
                if (hasLiveKernelDmp)
                {
                    powerData.LiveKernelDumpFlag = "(LiveKernel)";
                    powerData.LiveKernelCode = currLiveKernelDetails;
                }
                else
                {
                    powerData.LiveKernelDumpFlag = "";
                }

                if (liveKernelReportList.Count() > 0)
                {
                    powerData.LiveKernelList = liveKernelReportList;
                }
            }

            try
            {
                if (PSTDataProcessing.GetEnergyDrainRate(pstDeviceFolder, out double results))
                {
                    powerData.EnergyDrainRate = results;
                }
            }
            catch (Exception)
            {
                //ignore
            }

            if (PSTDataProcessing.GetCStatesActiveData(pstDeviceFolder, out double[] cStatesAvg))
            {
                if (cStatesAvg[0] != 0.0)
                {
                    powerData.C2ActivePercent = cStatesAvg[0];
                    reportData.C2ActiveAvgPer = true;
                }
                if (cStatesAvg[1] != 0.0)
                {
                    powerData.C3ActivePercent = cStatesAvg[1];
                    reportData.C3ActiveAvgPer = true;
                }
                if (cStatesAvg[2] != 0.0)
                {
                    powerData.C6ActivePercent = cStatesAvg[2];
                    reportData.C6ActiveAvgPer = true;
                }
                if (cStatesAvg[3] != 0.0)
                {
                    powerData.C7ActivePercent = cStatesAvg[3];
                    reportData.C7ActiveAvgPer = true;
                }
                if (cStatesAvg[4] != 0.0)
                {
                    powerData.C8ActivePercent = cStatesAvg[4];
                    reportData.C8ActiveAvgPer = true;
                }
                if (cStatesAvg[5] != 0.0)
                {
                    powerData.C9ActivePercent = cStatesAvg[5];
                    reportData.C9ActiveAvgPer = true;
                }
                if (cStatesAvg[6] != 0.0)
                {
                    powerData.C10ActivePercent = cStatesAvg[6];
                    reportData.C10ActiveAvgPer = true;
                }
                if (cStatesAvg[7] != 0.0)
                {
                    powerData.SleepS0Percent = cStatesAvg[7];
                    reportData.S0SleepAvgPer = true;
                }
            }
            return true;
        }

        private void GetPSTFolders(DirectoryInfo[] dirs, ref List<string> pstDirs)
        {
            foreach(DirectoryInfo di in dirs)
            {
                if (CheckForPSTResults(di.FullName))
                {
                    pstDirs.Add(di.FullName);
                }
                else
                {
                    GetPSTFolders(di.GetDirectories(), ref pstDirs);
                }
            }
        }

        private bool CheckForPSTResults(string dirName)
        {
            string deviceInfoFilePath = dirName + @"\DeviceInfo.csv";

            if (File.Exists(deviceInfoFilePath))
            {
                return true;
            }
            return false;
        }

        private void DisplayHelp()
        {
            Console.WriteLine("     PSTPowerReports [/h1=\"Header1\"] [/h2=\"Header2\"] [/PSTVer=\"2.3.1\"] /dir=\"c:\\TestResultFolders\" ");
            Console.WriteLine("         /? - displays help");
            Console.WriteLine("         /h1=Header1         - Adds description of the specific Run to the report");
            Console.WriteLine("         /h2=Header2         - Adds description of build version to the report");
            Console.WriteLine("         /PSTVer=\"2.3.1\"     - Adds PST version to the report");
            Console.WriteLine("         /dir=folderName     - The PST Testresults folder for generating report");
        }
    }
}
