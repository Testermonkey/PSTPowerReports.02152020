using System;
using System.IO;
using System.Linq;
using System.Collections.Generic;

namespace PSTPowerReports
{
    class ReportData
    {
        private string _header1;
        private string _header2;
        private string _PSTVer;
        private string _csTimeHeader;
        private int _csTime;
        private string _currPath;
        private bool _c2ActiveAvgPer;
        private bool _c3ActiveAvgPer;
        private bool _c6ActiveAvgPer;
        private bool _c7ActiveAvgPer;
        private bool _c8ActiveAvgPer;
        private bool _c9ActiveAvgPer;
        private bool _c10ActiveAvgPer;
        private bool _s0SleepAvgPer;
        private bool _hwDrips;
        private bool _swDrips;
        private bool _BatteryData;
        private bool _ActiveEnergyDrain;
        private bool _ShowActiveEnergy;

        public ReportData()
        {
            _header1 = string.Empty;
            _header2 = string.Empty;
            _PSTVer = string.Empty;
            _currPath = string.Empty;
            _c2ActiveAvgPer = false;
            _c3ActiveAvgPer = false;
            _c6ActiveAvgPer = false;
            _c7ActiveAvgPer = false;
            _c8ActiveAvgPer = false;
            _c9ActiveAvgPer = false;
            _c10ActiveAvgPer = false;
            _s0SleepAvgPer = false;
            _hwDrips = false;
            _swDrips = false;
            _BatteryData = false;
            _ActiveEnergyDrain = false;
            _ShowActiveEnergy = false;
    }


    public string Header1
        {
            set { _header1 = value; }
            get { return _header1; }
        }

        public string Header2
        {
            set { _header2 = value; }
            get { return _header2; }
        }

        public string PSTVer
        {
            set { _PSTVer = value; }
            get { return _PSTVer; }
        }

        public string CurrPath
        {
            set { _currPath = value; }
            get { return _currPath; }
        }

        public string CSTimeHeader
        {
            set { _csTimeHeader = value; }
            get { return _csTimeHeader; }
        }

        public bool C2ActiveAvgPer
        {
            set { _c2ActiveAvgPer = value; }
            get { return _c2ActiveAvgPer; }
        }

        public bool C3ActiveAvgPer
        {
            set { _c3ActiveAvgPer = value; }
            get { return _c3ActiveAvgPer; }
        }

        public bool C6ActiveAvgPer
        {
            set { _c6ActiveAvgPer = value; }
            get { return _c6ActiveAvgPer; }
        }

        public bool C7ActiveAvgPer
        {
            set { _c7ActiveAvgPer = value; }
            get { return _c7ActiveAvgPer; }
        }

        public bool C8ActiveAvgPer
        {
            set { _c8ActiveAvgPer = value; }
            get { return _c8ActiveAvgPer; }
        }

        public bool C9ActiveAvgPer
        {
            set { _c9ActiveAvgPer = value; }
            get { return _c9ActiveAvgPer; }
        }

        public bool C10ActiveAvgPer
        {
            set { _c10ActiveAvgPer = value; }
            get { return _c10ActiveAvgPer; }
        }

        public bool S0SleepAvgPer
        {
            set { _s0SleepAvgPer = value; }
            get { return _s0SleepAvgPer; }
        }
        public bool HwDrips
        {
            set { _hwDrips = value; }
            get { return _hwDrips; }
        }
        public bool SwDrips
        {
            set { _swDrips = value; }
            get { return _swDrips; }
        }
        public bool BatteryData
        {
            set { _BatteryData = value; }
            get { return _BatteryData; }
        }
        public bool ActiveEnergyDrain
        {
            set { _ActiveEnergyDrain = value; }
            get { return _ActiveEnergyDrain; }
        }
        public bool ShowActiveEnergy
        {
            set { _ShowActiveEnergy = value; }
            get { return _ShowActiveEnergy; }
        }

        public void Process(List<string> dirs)
        {
            if (dirs.Count() > 0)
            {
                if (!RetriveCSTime(dirs))
                {
                    _csTime = 0;
                    CSTimeHeader = "CS time not specified";
                }
                else
                {
                    CSTimeHeader = "CS time = " + _csTime / 60 + "min";
                }

                RetrivePSTVer(dirs);
            }
        }

        private bool RetriveCSTime(List<string> dirs)
        {
            foreach (string currDir in dirs)
            {
                string csTimeFile = currDir + @"\CStimeLog.csv";

                try
                {
                    if (File.Exists(csTimeFile))
                    {
                        using (StreamReader sr = new StreamReader(csTimeFile))
                        {
                            string line = string.Empty;
                            while ((line = sr.ReadLine()) != null)
                            {
                                if (!line.Contains("Session#"))
                                {
                                    string[] csTimeLogs = line.Split(',');
                                    if (csTimeLogs.Length >= 4)
                                    {
                                        if (int.TryParse(csTimeLogs[1].Trim(), out int csDuration))
                                        {
                                            _csTime = csDuration;
                                            return true;
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
                catch (Exception e)
                {
                    Util.Trace("Exception in ReportData::GetCSTime -> {0}", e.Message);
                    return false;
                }
            }

            return false;
        }

        private bool RetrivePSTVer(List<string> dirs)
        {
            foreach (string currDir in dirs)
            {
                string pstVerFile = currDir + @"\Trace\PSTtrace.txt";

                try
                {
                    if (File.Exists(pstVerFile))
                    {
                        using (StreamReader sr = new StreamReader(pstVerFile))
                        {
                            int cnt = 10;
                            string line = string.Empty;
                            while ((line = sr.ReadLine()) != null)
                            {
                                string pstStr = "power state stress test";
                                if (line.ToLower().Contains(pstStr) && line.ToLower().Trim().EndsWith("- start"))
                                {
                                    string verStr = line.Substring(line.ToLower().IndexOf(pstStr) + pstStr.Length).Trim();
                                    string[] verTokens = verStr.Split(' ');
                                    if (verTokens.Count() > 0)
                                    {
                                        if (verTokens[0].Trim().Contains("."))
                                        {
                                            _PSTVer = verTokens[0].Trim();
                                            return true;
                                        }
                                    }
                                }
                                if (--cnt <= 0)
                                    break;
                            }
                        }
                    }
                }
                catch (Exception e)
                {
                    Util.Trace("Exception in ReportData::RetrivePSTVer -> {0}", e.Message);
                    return false;
                }
            }

            return false;
        }
    }
}
