using System;
using System.IO;
using System.Xml;
using System.Xml.Serialization;
using System.Collections;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Bibliography;
using DocumentFormat.OpenXml.Office2016.Drawing.ChartDrawing;
using System.Threading;

namespace PSTPowerReports
{
    class PSTDataProcessing
    {
        static public bool GetDeviceDateAndName(string deviceFolder, out DateTime date, out string name)
        {
            name = string.Empty;
            date = DateTime.Now;

            if (!string.IsNullOrEmpty(deviceFolder))
            {
                int startIndex = deviceFolder.IndexOf("_20") + 1;
                name = deviceFolder.Substring(0, deviceFolder.IndexOf("_20"));
                if (startIndex <= 0)
                {
                    return false;
                }

                if (ParseToken(deviceFolder, startIndex, 4, out int deviceYear))
                {
                    if (SetYear(deviceYear, ref date))
                    {
                        startIndex += 5; // 4 for year length and 1 for '_'

                        if (ParseToken(deviceFolder, startIndex, 2, out int deviceMonth))
                        {
                            if (SetMonth(deviceMonth, ref date))
                            {
                                startIndex += 3; // 2 for month length and 1 for '-'

                                if (ParseToken(deviceFolder, startIndex, 2, out int deviceDay))
                                {
                                    if (SetDay(deviceDay, ref date))
                                    {
                                        startIndex += 3; // 2 for day length and 1 for '_'

                                        if (ParseToken(deviceFolder, startIndex, 2, out int deviceHour))
                                        {
                                            if (SetHour(deviceHour, ref date))
                                            {
                                                startIndex += 3; // 2 for hour length and 1 for '-'

                                                if (ParseToken(deviceFolder, startIndex, 2, out int deviceMinute))
                                                {
                                                    if (SetMinute(deviceMinute, ref date))
                                                    {
                                                        return true;
                                                    }
                                                }
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
            }
            return false;
        }

        static public bool GetOSType(string deviceFolder, out string osType)
        {
            string versionInfoFile = deviceFolder + @"\VersionInformation.txt";
            int buildNum = 0;
            osType = Win10OSType.None.ToString();

            try
            {
                if (File.Exists(versionInfoFile))
                {
                    using (StreamReader sr = new StreamReader(versionInfoFile))
                    {
                        string line = string.Empty;
                        while ((line = sr.ReadLine()) != null)
                        {
                            if (line.Contains("CurrentBuild"))
                            {
                                string currBuild = line.Substring((line.IndexOf("CurrentBuild") + "CurrentBuild".Length));
                                if (!int.TryParse(currBuild.Trim(), out buildNum))
                                {
                                    buildNum = 0;
                                }
                                break;
                            }
                        }
                        sr.Close();
                    }

                    osType = GetWin10OSType(buildNum);
                }
                return true;
            }
            catch (Exception e)
            {
                Util.Trace("Exception in PSTDataProcessing::GetOSType -> {0}", e.Message);
                return false;
            }
        }

        static public bool GetPowerSliderType(string deviceFolder, out PowerSliderType pwrMode)
        {
            pwrMode = PowerSliderType.None;
            bool retStatus = false;
            string powerCfgOverlayFile = deviceFolder + @"\Power\powercfg_overlay.txt";
            string powerCfgFile = deviceFolder + @"\Power\powercfg.txt";

            try
            {
                if (File.Exists(powerCfgOverlayFile))
                {
                    using (StreamReader sr = new StreamReader(powerCfgOverlayFile))
                    {
                        string line = string.Empty;
                        while ((line = sr.ReadLine()) != null)
                        {
                            if (line.Contains("Max Performance Overlay"))
                            {
                                pwrMode = PowerSliderType.Best;
                                retStatus = true;
                                break;
                            }

                            if (line.Contains("High Performance Overlay"))
                            {
                                pwrMode = PowerSliderType.Better;
                                retStatus = true;
                                break;
                            }
                        }
                    }
                }

                if ((File.Exists(powerCfgFile)) && (pwrMode == PowerSliderType.None))
                {
                    using (StreamReader sr = new StreamReader(powerCfgFile))
                    {
                        string line = string.Empty;
                        bool perfeppFlag = false;

                        while ((line = sr.ReadLine()) != null)
                        {
                            if (!perfeppFlag)
                            {
                                if (line.Contains("Processor energy performance preference policy"))
                                {
                                    perfeppFlag = true;
                                }
                            }
                            else
                            {
                                if (line.Contains("Current DC Power Setting Index"))
                                {
                                    string[] dcValues = line.Split(':');
                                    if (dcValues.Length >= 2)
                                    {
                                        string dcValue = dcValues[1].Trim();
                                        if (UInt64.TryParse(dcValue.Substring(dcValue.IndexOf("0x") + "0x".Length),
                                            NumberStyles.HexNumber, CultureInfo.CurrentCulture, out ulong currDCValue))
                                        {
                                            if (currDCValue == 70)
                                            {
                                                pwrMode = PowerSliderType.Saver;
                                                retStatus = true;
                                                break;
                                            }
                                        }
                                    }
                                }
                                else if (string.IsNullOrEmpty(line))
                                { // reached end of PEREPP setting
                                    pwrMode = PowerSliderType.None;
                                    retStatus = true;
                                    break;
                                }
                            }
                        }
                    }
                }
            }
            catch (Exception e)
            {
                Util.Trace("Exception in PSTDataProcessing::GetPowerSliderType -> {0}", e.Message);
                retStatus = false;
            }

            return retStatus;
        }

        static public bool GetPowerSliderTypeFromDevice(string deviceFolder, out bool isAC, out PowerSliderType pwrMode, out double totalBatLevel,
            out double bat1Level, out double bat2Level, out string netType)
        {
            pwrMode = PowerSliderType.None;
            bat1Level = 0.0;
            bat2Level = 0.0;
            totalBatLevel = 0.0;
            isAC = false;
            netType = string.Empty;
            string powerSliderStatusFile = string.Empty;
            bool retStatus = false; //"PowerSliderState.CSV  or PowerStatusLog*.csv

            if (File.Exists(deviceFolder + @"\PowerStatusLog.csv"))
            {
                powerSliderStatusFile = deviceFolder + @"\PowerStatusLog.csv";
            }
            else if (File.Exists(deviceFolder + @"\PowerSliderStatus.csv"))
            {
                powerSliderStatusFile = deviceFolder + @"\PowerSliderStatus.csv";
            }
            else if (File.Exists(deviceFolder + @"\PowerSliderState.csv"))
            {
                powerSliderStatusFile = deviceFolder + @"\PowerSliderState.csv";
            }

            try
            {
                if (File.Exists(powerSliderStatusFile))
                {
                    Console.WriteLine(powerSliderStatusFile);
                    string powerSliderStatusRecord = string.Empty;
                    string lastPowerSliderStatusRecord = string.Empty;
                    string[] psStatusFileList = File.ReadAllLines(powerSliderStatusFile);

                    if (psStatusFileList.Count() > 0)
                    {
                        if (psStatusFileList.Any(x => x.ToUpper().Contains("FINAL")))
                        {
                            //powerSliderStatusRecord = psStatusFileList.First(x => x.ToUpper().Contains("FINAL"));
                            Console.WriteLine(psStatusFileList.First(x => x.ToUpper().Contains("FINAL")));
                            int recordcounter = (psStatusFileList.Count() - 2);
                            powerSliderStatusRecord = psStatusFileList[recordcounter];
                        }
                        else if (psStatusFileList.Any(x => x.ToUpper().Contains("INITIAL")))
                        {
                            Console.WriteLine(psStatusFileList.First(x => x.ToUpper().Contains("INITIAL")));
                            powerSliderStatusRecord = psStatusFileList.First(x => x.ToUpper().Contains("INITIAL"));
                        }
                        else if (psStatusFileList[0].Contains("Session"))
                        {
                            if (psStatusFileList.Count() >= 2)
                            {
                                powerSliderStatusRecord = psStatusFileList[1];
                                lastPowerSliderStatusRecord = psStatusFileList[(psStatusFileList.Count() - 1)];
                            }
                        }
                        if (!string.IsNullOrEmpty(powerSliderStatusRecord))
                        {
                            string powerSliderState = string.Empty;
                            string powerStateOfDevice = string.Empty;
                            string PowerMode = string.Empty;

                            // get the last line to use for battery levels
                            int lineitem = psStatusFileList.Count() - 1;
                            string line = psStatusFileList[lineitem];
                            //string[] list = line.Split(',');

                            // holds a record for use in the method
                            string[] pssToken = powerSliderStatusRecord.Split(',');
                            // Holds the top line header values to determine the column to use in the method
                            string[] pssColumnToken = psStatusFileList[0].Split(',');
                            for (int ii = 0; ii < pssColumnToken.Count(); ii++)
                            {
                                if (pssColumnToken[ii].Trim().ToLower().Equals("power") || pssColumnToken[ii].Trim().ToLower().Equals("powertype")) //NetInterface
                                {
                                    powerStateOfDevice = pssToken[ii].Trim();
                                }
                                else if (pssColumnToken[ii].Trim().ToLower().Equals("slider") || pssColumnToken[ii].Trim().ToLower().Equals("powerslider"))
                                {
                                    powerSliderState = pssToken[ii].Trim();
                                }
                                else if (pssColumnToken[ii].Trim().ToLower().Equals("netinterface"))
                                {
                                    netType = pssToken[ii].Trim();
                                }
                                else if (pssColumnToken[ii].Trim().ToLower().Equals("batterylevel"))
                                {
                                    try
                                    {
                                        totalBatLevel = Convert.ToDouble(pssToken[ii]);
                                    }
                                    catch (Exception)
                                    { }
                                }
                                else if (pssColumnToken[ii].Trim().ToLower().Equals("bat01"))
                                {
                                    try
                                    {
                                        bat1Level = Convert.ToDouble(pssToken[ii]);
                                    }
                                    catch (Exception)
                                    {
                                    }
                                }
                                else if (pssColumnToken[ii].Trim().ToLower().Equals("bat02"))
                                {
                                    try
                                    {
                                        bat2Level = Convert.ToDouble(pssToken[ii]);
                                    }
                                    catch (Exception)
                                    {
                                    }
                                }
                                else if (pssColumnToken[ii].Trim().ToLower().Equals("power mode"))  // oldest method
                                {

                                    string[] modevalue = null;
                                    try
                                    {
                                        modevalue = pssToken[ii].Split(':');
                                        powerStateOfDevice = modevalue[0].Trim();
                                        powerSliderState = modevalue[1].Trim();
                                    }
                                    catch (Exception)
                                    {
                                        powerStateOfDevice = "dc";
                                        powerSliderState = "RECOMMENDED";
                                    }
                                }
                                else if (pssColumnToken[ii].Trim().ToLower().Equals("powerstate"))  // oldest method
                                {
                                     PowerMode = pssToken[ii].Trim(); 

                                }
                            }

                            if (!string.IsNullOrEmpty(powerStateOfDevice))
                            {
                                if (powerStateOfDevice.ToUpper().Contains("AC") || powerStateOfDevice.ToLowerInvariant().Contains("plugged in"))
                                {
                                    isAC = true;
                                }
                                else if (powerStateOfDevice.ToUpper().Contains("DC") || powerStateOfDevice.ToLowerInvariant().Contains("on battery") )
                                {
                                    isAC = false;
                                }
                                else
                                { // could not find "on battery" OR "plugged in" for Power mode
                                    return false;
                                }

                                if (!string.IsNullOrEmpty(powerSliderState))
                                {
                                    if (powerSliderState.ToUpper().Contains("BEST"))
                                    {
                                        pwrMode = PowerSliderType.Best;
                                        retStatus = true;
                                    }
                                    else if (powerSliderState.ToUpper().Contains("BETTER"))
                                    {
                                        pwrMode = PowerSliderType.Better;
                                        retStatus = true;
                                    }
                                    else if (powerSliderState.ToUpper().Contains("RECOMMENDED"))
                                    {
                                        pwrMode = PowerSliderType.Recommended;
                                        retStatus = true;
                                    }
                                    else if (powerSliderState.ToUpper().Contains("SAVER"))
                                    {
                                        pwrMode = PowerSliderType.Saver;
                                        retStatus = true;
                                    }
                                }
                                else
                                {
                                    pwrMode = PowerSliderType.Recommended;
                                    retStatus = true;
                                }
                            }
                        }
                    }
                    // battery data collection from last line of report?
                }
            }
            catch (Exception e)
            {
                Util.Trace("Exception in PSTDataProcessing::GetPowerSliderTypeFromDevice -> {0}", e.Message);
                retStatus = false;
            }

            return retStatus;
        }

        static public bool CheckMemoryDump(string deviceFolder, out bool hasMemDump, out List<string> bugcheckDetails, out List<string> memDumpList)
        {
            char[] _trim_hex = new char[] { '0', 'x', 'X', ' ' };
            string crashReportFile = deviceFolder + @"\CrashReport.csv";
            string memDumpDir = deviceFolder + @"\MemDump";
            hasMemDump = false;
            bugcheckDetails = new List<string>();
            memDumpList = new List<string>();

            try
            {
                if (File.Exists(crashReportFile))
                {
                    using (StreamReader sr = new StreamReader(crashReportFile))
                    {
                        string line = string.Empty;
                        while ((line = sr.ReadLine()) != null)
                        {
                            if (line.Contains("CRASHED,"))
                            {
                                hasMemDump = true;
                            }
                            if ((!line.Contains("Session#,")) && (line.Contains("CRASHED,")))
                            {
                                string[] crashDetails = line.Split(',');
                                if (crashDetails.Length >= 8)
                                {
                                    string currCrash = crashDetails[3] + " {";
                                    if (ulong.TryParse(crashDetails[4].Trim().TrimStart(_trim_hex), NumberStyles.HexNumber, null, out ulong param1))
                                    {
                                        currCrash += param1.ToString("x") + ", ";
                                    }
                                    if (ulong.TryParse(crashDetails[5].Trim().TrimStart(_trim_hex), NumberStyles.HexNumber, null, out ulong param2))
                                    {
                                        currCrash += param2.ToString("x") + ", ";
                                    }
                                    if (ulong.TryParse(crashDetails[6].Trim().TrimStart(_trim_hex), NumberStyles.HexNumber, null, out ulong param3))
                                    {
                                        currCrash += param3.ToString("x") + ", ";
                                    }
                                    if (ulong.TryParse(crashDetails[7].Trim().TrimStart(_trim_hex), NumberStyles.HexNumber, null, out ulong param4))
                                    {
                                        currCrash += param4.ToString("x") + "}";
                                    }
                                    if (!string.IsNullOrEmpty(crashDetails[8].Trim()))
                                    {
                                        currCrash += " " + crashDetails[8].Trim().ToUpper();
                                    }
                                    bugcheckDetails.Add(currCrash);
                                }
                            }
                        }
                        sr.Close();
                    }
                }
                if (Directory.Exists(memDumpDir))
                {
                    memDumpList = Directory.GetFiles(memDumpDir).ToList<string>();
                }
                return true;
            }
            catch (Exception e)
            {
                Util.Trace("Exception in PSTDataProcessing::CheckMemoryDump -> {0}", e.Message);
                return false;
            }
        }

        static public bool CheckLiveKernelDump(string deviceFolder, out bool hasLiveKernelDump, out List<string> liveKernelDetails, out List<string> liveKernelList)
        {
            char[] _trim_hex = new char[] { '0', 'x', 'X', ' ' };
            string liveKernelReportFile = deviceFolder + @"\LiveKernelReport.csv";
            string liveKernelDir = deviceFolder + @"\LiveKernelReports";
            hasLiveKernelDump = false;
            liveKernelDetails = new List<string>();
            liveKernelList = new List<string>();

            try
            {
                if (File.Exists(liveKernelReportFile))
                {
                    using (StreamReader sr = new StreamReader(liveKernelReportFile))
                    {
                        string line = string.Empty;
                        while ((line = sr.ReadLine()) != null)
                        {
                            if (!line.Contains("Serial#"))
                            {
                                string[] lkrDetails = line.Split(',');
                                if (lkrDetails.Length >= 8)
                                {
                                    string currLiveKernelReport = lkrDetails[2] + " {";
                                    if (ulong.TryParse(lkrDetails[3].Trim().TrimStart(_trim_hex), NumberStyles.HexNumber, null, out ulong param1))
                                    {
                                        currLiveKernelReport += param1.ToString("x") + ", ";
                                    }
                                    if (ulong.TryParse(lkrDetails[4].Trim().TrimStart(_trim_hex), NumberStyles.HexNumber, null, out ulong param2))
                                    {
                                        currLiveKernelReport += param2.ToString("x") + ", ";
                                    }
                                    if (ulong.TryParse(lkrDetails[5].Trim().TrimStart(_trim_hex), NumberStyles.HexNumber, null, out ulong param3))
                                    {
                                        currLiveKernelReport += param3.ToString("x") + ", ";
                                    }
                                    if (ulong.TryParse(lkrDetails[6].Trim().TrimStart(_trim_hex), NumberStyles.HexNumber, null, out ulong param4))
                                    {
                                        currLiveKernelReport += param4.ToString("x") + "}";
                                    }
                                    if (!string.IsNullOrEmpty(lkrDetails[7].Trim()))
                                    {
                                        currLiveKernelReport += " " + lkrDetails[7].Trim();
                                    }
                                    liveKernelDetails.Add(currLiveKernelReport);
                                    hasLiveKernelDump = true;
                                }
                            }
                        }
                        sr.Close();
                    }
                }
                if (Directory.Exists(liveKernelDir))
                {
                    try
                    {
                        DirectoryInfo dir = new DirectoryInfo(liveKernelDir);
                        DirectoryInfo[] dirs = dir.GetDirectories();
                        FileInfo[] lkFileInfos = dir.GetFiles();

                        if (lkFileInfos.Count() > 0)
                        {
                            liveKernelList = Directory.GetFiles(liveKernelDir).ToList<string>();
                        }

                        if (dirs.Count() > 0)
                        {
                            foreach (DirectoryInfo di in dirs)
                            {
                                FileInfo[] diFiles = di.GetFiles();
                                foreach (FileInfo dfi in diFiles)
                                {
                                    FileInfo lkFile = lkFileInfos.FirstOrDefault(x => dfi.Name.Contains(x.Name));
                                    if (lkFile != null)
                                    {
                                        if (dfi.Length > lkFile.Length)
                                        {
                                            liveKernelList.Remove(lkFile.FullName);
                                            liveKernelList.Add(dfi.FullName);
                                        }
                                    }
                                    else
                                    {
                                        liveKernelList.Add(dfi.FullName);
                                    }
                                }
                            }
                        }
                    }
                    catch (Exception e)
                    {
                        Util.Trace("Exception in PSTDataProcessing::CheckLiveKernelDump (extracting live kernel reports) -> {0}", e.Message);
                        return false;
                    }
                }
            }
            catch (Exception e)
            {
                Util.Trace("Exception in PSTDataProcessing::CheckLiveKernelDump -> {0}", e.Message);
                return false;
            }
            return true;
        }

        static public bool GetEnergyDrainRate(string deviceFolder, out double avgDrainRate)
        {
            try
            {
                string energyDrainFile = deviceFolder + @"\Power\EnergyDrain.csv";
                string altenergyDrainFile = deviceFolder + @"\EnergyDrain.csv";
                bool foundEnergydrains = false;
                if (!File.Exists(deviceFolder + @"\Power\EnergyDrain.csv"))
                {
                    if (File.Exists(deviceFolder + @"\EnergyDrain.csv"))
                    {
                        energyDrainFile = altenergyDrainFile; 
                        foundEnergydrains = true;
                    }
                }
                else
                {
                    foundEnergydrains = true;
                }
                avgDrainRate = 0.0;
                if (foundEnergydrains)
                {
                    if (GetEnegyDrainAvgData(energyDrainFile, 30, 1900, out double result))
                    {
                        avgDrainRate = result; //SL3 30 to 1900, previous 70, 1300
                        return true;
                    }
                }
            }
            catch (Exception e)
            {
                avgDrainRate = 0;
                  Util.Trace("Exception in PSTDataProcessing::GetEnergyDrainRate -> {0}", e.Message);
                return false;
            }
            return false;
        }


        // New feature adds Drips data and Pulls Active energy drain rates from sleep study.xml for each run. not recursive...
        // writes a DrainRate.xml to disk
        static public bool GetDrips(string deviceFolder, out double HwDripRate, out double SwDripRate, out double ActiveEnergyDrainRate, bool showActiveEnergy)
        {
            string studyFile = deviceFolder + @"\Power\sleepstudy-report_verbose.xml";
            string altstudyFile = deviceFolder + @"\sleepstudy.xml";
            string templocation = deviceFolder + @"\Power";
            string activeDrains = deviceFolder + @"\Power\ActiveDrains.csv";
            string altactiveDrains = deviceFolder + @"\ActiveDrains.csv";
            HwDripRate = 0.0;
            SwDripRate = 0.0;
            ActiveEnergyDrainRate = 0.0;
            double TotalDuration = 0.0;
            
            try
            {
                // Some older test plans are not setting the correct stuff, so will try to use the other sleep study xml from the root folder.
                bool FileFound = false;
                if (File.Exists(studyFile))
                {
                    FileFound = true;
                }
                else if (File.Exists(altstudyFile))
                {
                    FileFound = true;
                    studyFile = altstudyFile;
                }

                if (FileFound)
                {
                    if (GetSleepReportHwDrainData(studyFile, out List<double> swDripResult, out List<double> hwDripResult, out List<double> durationTicks, out List<powerConsumptionData> activeDrainRateData))
                    {
                        for (int i = 0; i < swDripResult.Count; i++)
                        {
                            try
                            {
                                HwDripRate += hwDripResult[i];
                                SwDripRate += swDripResult[i];
                                TotalDuration += durationTicks[i];
                            }
                            catch (Exception)
                            {
                            }
                        }
                            // now get the total avg?
                            HwDripRate = (HwDripRate / TotalDuration) * 100;
                            SwDripRate = (SwDripRate / TotalDuration) * 100;
                        // now to round the numbers for a percentage?
                        HwDripRate = Math.Round(HwDripRate,3);
                        SwDripRate = Math.Round(SwDripRate,3);


                        // now for the active drainrate calculations
                        double mwh = 0.0;
                        TimeSpan ts = TimeSpan.Zero;
                        // prep to write to XML file

                        bool DrainsFileFound = false;
                        if (Directory.Exists(templocation))
                        {
                            DrainsFileFound = true;
                        }
                        else if (Directory.Exists(deviceFolder))
                        {
                            DrainsFileFound = true;
                            activeDrains = altactiveDrains;
                        }

                        if (DrainsFileFound)
                        {
                            try
                            {
                                using (FileStream fs = File.Create(activeDrains))
                                {
                                }
                            }
                            catch (Exception ex)
                            {
                                Console.WriteLine("Failed to create the file for active drains");
                                Console.WriteLine(ex.ToString());
                            }
                        }
                        using (StreamWriter wr = new StreamWriter(activeDrains, true))
                        {
                            wr.WriteLine("{0},{1},{2},{3}", "Drain_Start", "Drain_End", "Start_Capacity", "End_Capacity");
                            foreach (var item in activeDrainRateData)
                            {
                                try
                                {
                                    wr.WriteLine(String.Format("{0},{1},{2},{3}", item.DtStart, item.DtEnd, item.StartChargeCap, item.EndChargeCap));
                                }
                                catch (Exception)
                                {
                                    Console.WriteLine("Failed to write a line to the file for active drains");
                                }
                                TimeSpan tempspan = new TimeSpan(item.DtEnd.Ticks - item.DtStart.Ticks);
                                ts += tempspan;
                                double tempChargeDrained = item.StartChargeCap - item.EndChargeCap;
                                mwh += tempChargeDrained;
                            }
                        }
                        
                        // take the total battery drain in mWh and divide by hours..  be sure to convert each value to corresponding values for time and mWh...
                        int tempTimeinhours = ts.Hours;
                        if (showActiveEnergy)
                        {
                            ActiveEnergyDrainRate = mwh / tempTimeinhours;
                        }
                        else
                        {
                            ActiveEnergyDrainRate = 0;
                        }
                    }
                }
            }
            catch (Exception e)
            {
                Util.Trace("Exception in PSTDataProcessing::GetEnergyDrainRate -> {0}", e.Message);
                return false;
            }
            return true;
        }

        static private bool WriteDrainRateDatatoXML(List<powerConsumptionData> activeDrainRateData) 
        {


            return true;
        }

        static public bool GetCStatesActiveData(string deviceFolder, out double[] cStatesActivePercent)
        {
            int count = 0;
            string cStatesFile = deviceFolder + @"\Power\CStateInfo.csv";
            cStatesActivePercent = new double[8];

            //initialize data
            for (int i = 0; i < 8; i++)
            {
                cStatesActivePercent[i] = 0.0;
            }

            try
            {
                if (File.Exists(cStatesFile))
                {
                    using (StreamReader sr = new StreamReader(cStatesFile))
                    {
                        string line = string.Empty;
                        while ((line = sr.ReadLine()) != null)
                        {
                            if (!line.Contains("ActivePer")) // ignore the file header
                            {
                                string[] cStatesDetails = line.Split(',');
                                if (cStatesDetails.Length >= 8)
                                {
                                    for (int i = 0; i < 8; i++)
                                    {
                                        if (double.TryParse(cStatesDetails[i].Trim('"'), out double cStatesResult))
                                        {
                                            cStatesActivePercent[i] += cStatesResult;
                                        }
                                    }
                                    count++;
                                }
                            }
                        }
                    }

                    for (int i = 0; i < 8; i++)
                    {
                        if ((cStatesActivePercent[i] == 0.0) || (count == 0))
                        {
                            cStatesActivePercent[i] = 0.0;
                        }
                        else
                        {
                            cStatesActivePercent[i] /= count;
                        }
                    }
                }
                return true;
            }
            catch (Exception e)
            {
                Util.Trace("Exception in PSTDataProcessing::GetCStatesActiveData -> {0}", e.Message);
                return false;
            }
        }

        static public double GetSleepS0Percent(string deviceFolder)
        {
            return 0.0;
        }

        static private bool ParseToken(string devFolder, int startIdx, int len, out int tokenValue)
        {
            if (devFolder.Length >= (startIdx + len))
            {
                if (int.TryParse(devFolder.Substring(startIdx, len), out tokenValue))
                {
                    return true;
                }
            }
            tokenValue = 0;
            return false;
        }

        static private bool SetYear(int year, ref DateTime dt)
        {
            try
            {
                if ((year >= 2000) && (year <= 2099))
                {
                    int deltaYear = year - dt.Year;
                    dt = dt.AddYears(deltaYear);
                    return true;
                }
            }
            catch
            {
                return false;
            }

            return false;
        }

        static private bool SetMonth(int month, ref DateTime dt)
        {
            try
            {
                if ((month >= 1) && (month <= 12))
                {
                    int deltaMonth = month - dt.Month;
                    dt = dt.AddMonths(deltaMonth);
                    return true;
                }
            }
            catch
            {
                return false;
            }

            return false;
        }

        static private bool SetDay(int day, ref DateTime dt)
        {
            try
            {
                if ((day >= 1) && (day <= 31))
                {
                    int deltaDay = day - dt.Day;
                    dt = dt.AddDays(deltaDay);
                    return true;
                }
            }
            catch
            {
                return false;
            }

            return false;
        }

        static private bool SetHour(int hour, ref DateTime dt)
        {
            try
            {
                if ((hour >= 0) && (hour <= 24))
                {
                    int deltaHour = hour - dt.Hour;
                    dt = dt.AddHours(deltaHour);
                    return true;
                }
            }
            catch
            {
                return false;
            }

            return false;
        }

        static private bool SetMinute(int minute, ref DateTime dt)
        {
            try
            {
                if ((minute >= 0) && (minute <= 60))
                {
                    int deltaMinute = minute - dt.Minute;
                    dt = dt.AddMinutes(deltaMinute);
                    int deltaSecond = dt.Second;
                    dt = dt.AddSeconds((deltaSecond * -1));
                    return true;
                }
            }
            catch
            {
                return false;
            }

            return false;
        }

        static private string GetWin10OSType(int currentBuild)
        {
            PowerData wt = new PowerData();
            string currOSType = wt.GetDescription(Win10OSType.None);

            if (currentBuild == 10586)
            {
                currOSType = wt.GetDescription(Win10OSType._TH2);
            }
            else if (currentBuild == 14393)
            {
                currOSType = wt.GetDescription(Win10OSType._RS1);
            }
            else if ((currentBuild >= 15009) && (currentBuild < 16232))
            {
                currOSType = wt.GetDescription(Win10OSType._RS2);
            }
            else if ((currentBuild >= 16232) && (currentBuild <= 17000))
            {
                currOSType = wt.GetDescription(Win10OSType._RS3);
            }
            else if ((currentBuild >= 17001) && (currentBuild <= 17711))
            {
                currOSType = wt.GetDescription(Win10OSType._RS4);
            }
            else if ((currentBuild > 17711) && (currentBuild <= 18361))
            {
                currOSType = wt.GetDescription(Win10OSType._RS5);
            }
            else if ((currentBuild > 18361) && (currentBuild <= 18362))
            {
                currOSType = wt.GetDescription(Win10OSType._19H1);
            }
            else if ((currentBuild > 18362) && (currentBuild <= 18363))
            {
                currOSType = wt.GetDescription(Win10OSType._19H2);
            }
            else if ((currentBuild > 18363) && (currentBuild <= 19041))
            {
                currOSType = wt.GetDescription(Win10OSType._20H1);
            }
            else if ((currentBuild > 19041) && (currentBuild <= 19042))
            {
                currOSType = wt.GetDescription(Win10OSType._20H2);
            }
            else if ((currentBuild > 19042) && (currentBuild <= 19200))
            {
                currOSType = wt.GetDescription(Win10OSType._21H1);
            }
            else if ((currentBuild > 19200) && (currentBuild <= 19344))
            {
                currOSType = wt.GetDescription(Win10OSType._21H2);
            }
            else if ((currentBuild > 19344) && (currentBuild <= 2200))
            {
                currOSType = wt.GetDescription(Win10OSType._SunValley_Next);
            }
            else if (currentBuild > 22000) 
            {
                currOSType = wt.GetDescription(Win10OSType.FE);
            }
            else
            {
                currOSType = wt.GetDescription(Win10OSType.None);
            }

            return currOSType;
        }

        static private bool GetEnegyDrainAvgData(string fileName, int minValue, int maxValue, out double result)
        {
            int count = 0;
            result = 0.0;

            try
            {
                using (StreamReader sr = new StreamReader(fileName))
                {
                    string line = string.Empty;

                    while ((line = sr.ReadLine()) != null)
                    {
                        if (!line.Contains("EnergyDrain"))
                        {
                            string[] energyDrain = line.Split(',');
                            if (energyDrain.Length >= 1)
                            {
                                if (double.TryParse(energyDrain[0].Trim('"'), out double value))
                                {
                                    if ((value >= (double)minValue) && (value <= (double)maxValue))
                                    {
                                        result += value;
                                        count++;
                                    }
                                }
                            }
                        }
                    }
                    sr.Close();
                }
                if ((result == 0.0) || (count == 0))
                {
                    result = 0.0;
                }
                else
                {
                    result /= count;
                }

                return true;
            }
            catch (Exception e)
            {
                Util.Trace("Exception in PSTDataProcessing::GetEnegyDrainAvgData -> {0}", e.Message);
                return false;
            }
        }

        public partial class powerConsumptionData
        {
            private DateTime dtStart;
            private DateTime dtEnd;
            private double startChargeCap;
            private double endChargeCap;
            public powerConsumptionData() { }
            public powerConsumptionData(double startCap, double endCap,  DateTime startDT, DateTime endDT)
            {
                this.startChargeCap = startCap;
                this.endChargeCap = endCap;
                this.dtStart = startDT;
                this.dtEnd = endDT;
            }
            public DateTime DtStart
            {
                get { return dtStart; }
                set { dtStart = value; }
            }
            public DateTime DtEnd
            {
                get { return dtEnd; }
                set { dtEnd = value; }
            }
            public double StartChargeCap
            {
                get { return startChargeCap; }
                set { startChargeCap = value; }
            }
            public double EndChargeCap
            {
                get { return endChargeCap; }
                set { endChargeCap = value; }
            }
        }

        static private bool GetSleepReportHwDrainData(string fileName, out List<double> SwResult, out List<double> HwResult, out List<double> DurationResult, out List<powerConsumptionData> PowerCosmuptionData)
        {

            SwResult = new List<double>();
            HwResult = new List<double>();
            DurationResult = new List<double>();
            PowerCosmuptionData = new List<powerConsumptionData>();

            try
            {
             //   var list = nodeList.  ("ScenarioInstances");
                FileStream fs = new FileStream(fileName, FileMode.Open, FileAccess.Read);
                XmlDocument xmldoc = new XmlDocument();
                XmlNodeList xmlnode;
                XmlNodeList xmlDrainsnode;
                // now loading the xml data
                xmldoc.Load(fs);
                xmlnode = xmldoc.GetElementsByTagName("ScenarioInstance");
                xmlDrainsnode = xmldoc.GetElementsByTagName("EnergyDrains");


                string lpst = "";
                string hwpst = "";
                string duration = "";

                foreach (XmlNode drain in xmlDrainsnode)
                {
                    foreach (XmlNode childN in drain.ChildNodes)
                    {
                        for (int i = 0; i < childN.Attributes.Count; i++)
                        {
                            if (childN.Attributes[i].Name.ToLowerInvariant().Equals("ac"))
                            {
                                // now we have an active line, check for anything else or just grab the time and power comsumed
                                if (childN.Attributes[i].Value.Equals("0"))  // zero is discharging or NOT on A\C power.  One means charging.
                                {
                                    try
                                    {
                                        powerConsumptionData pcd = new powerConsumptionData();
                                        pcd.DtStart = DateTime.Parse(childN.Attributes[1].Value);
                                        pcd.DtEnd = DateTime.Parse(childN.Attributes[3].Value);
                                        pcd.StartChargeCap = Convert.ToDouble(childN.Attributes[4].Value);
                                        pcd.EndChargeCap = Convert.ToDouble(childN.Attributes[6].Value);
                                        PowerCosmuptionData.Add(pcd);
                                        break;
                                    }
                                    catch (Exception)
                                    {                                        
                                    }                                   
                                }
                            }
                        }
                    }
                }

                foreach (XmlNode node in xmlnode)
                {
                    foreach (XmlAttribute item in node.Attributes)
                    {
                        if (item.Name.Equals("LowPowerStateTime"))
                        {
                            lpst = item.Value;
                            try
                            {
                                SwResult.Add(Convert.ToDouble(lpst));
                            }
                            catch (Exception)
                            {
                                //   throw;
                            }
                        }
                        else if (item.Name.Equals("HwLowPowerStateTime"))
                        {
                            hwpst = item.Value;
                            try
                            {
                                HwResult.Add(Convert.ToDouble(hwpst));
                            }
                            catch (Exception)
                            {
                                throw;
                            }
                        }
                        else if (item.Name.Equals("Duration"))
                        {
                            duration = item.Value;
                            try
                            {
                                DurationResult.Add(Convert.ToDouble(duration));
                            }
                            catch (Exception)
                            {
                                throw;
                            }
                        }
                    }
                }
                return true;
            }
            catch (Exception e)
            {
                Util.Trace("Exception in PSTDataProcessing::GetEnegyDrainAvgData -> {0}", e.Message);
                return false;
            }
        }
    }
}

