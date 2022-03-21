using System;
using System.Collections.Generic;
using System.Linq;
using System.ComponentModel;
using System.Threading.Tasks;
using System.Reflection;

namespace PSTPowerReports
{
    public enum Win10OSType
    {
        [Description("None")]
        None = 0,
        [Description("TH2")]
        _TH2 = 1,
        [Description("RS1")]
        _RS1 = 2,
        [Description("RS2")]
        _RS2 = 3,
        [Description("RS3")]
        _RS3 = 4,
        [Description("RS4")]
        _RS4 = 5,
        [Description("RS5")]
        _RS5 = 6,
        [Description("19H1")]
        _19H1 = 7,
        [Description("19H2")]
        _19H2 = 8,
        [Description("20H1")]
        _20H1 = 9,
        [Description("20H2")]
        _20H2 = 10,
        [Description("21H1")]
        _21H1 = 11,
        [Description("21H2")]
        _21H2 = 12,
        [Description("SunValley_Next")]
        _SunValley_Next = 13,
        [Description("FE")]
        FE = 14
        // Win10OSType._TH2.GetDescription()
    }

    public enum PowerSliderType
    {
        None = 0,
        Best = 1,
        Better = 2,
        Recommended = 3,
        Saver = 4
    }

    public enum PowerSourceType
    {
        None = 0,
        Battery = 1,
        AC = 2
    }

    class PowerData
    {
        private DateTime _date;
        private string _deviceName;
        private string _relativePath;
        private string _osType;
        private PowerSliderType _powerSliderMode;
        private PowerSourceType _powerSourceType;
        private bool _hasMemoryDump;
        private string _memoryDumpFlag;
        private List<string> _bugCheckCode;
        private List<string> _memoryDumpList;
        private bool _hasLiveKernelDump;
        private bool _AdminReport;
        private string _liveKernelDumpFlag;
        private List<string> _liveKernelCode;
        private List<string> _liveKernelList;
        private double _energyDrainRate;
        private double _c2ActivePercent;
        private double _c3ActivePercent;
        private double _c6ActivePercent;
        private double _c7ActivePercent;
        private double _c8ActivePercent;
        private double _c9ActivePercent;
        private double _c10ActivePercent;
        private double _sleepS0Percent;
        private double _swDrips;
        private double _hwDrips;
        private double _TotalBatt;
        private double _Batt1;
        private double _Batt2;
        private double _ActiveEnergyDrainRate;
        private string _NetType;

        public PowerData()
        {
            _date = DateTime.Now;
            _deviceName = string.Empty;
            _relativePath = string.Empty;
            _osType = Win10OSType.None.ToString();
            _powerSliderMode = PowerSliderType.None;
            _powerSourceType = PowerSourceType.None;
            _hasMemoryDump = false;
            _memoryDumpFlag = string.Empty;
            _bugCheckCode = new List<string>();
            _memoryDumpList = new List<string>();
            _hasLiveKernelDump = false;
            _liveKernelDumpFlag = string.Empty;
            _liveKernelCode = new List<string>();
            _liveKernelList = new List<string>();
            _energyDrainRate = 0.0;
            _c2ActivePercent = 0.0;
            _c3ActivePercent = 0.0;
            _c6ActivePercent = 0.0;
            _c7ActivePercent = 0.0;
            _c8ActivePercent = 0.0;
            _c9ActivePercent = 0.0;
            _c10ActivePercent = 0.0;
            _sleepS0Percent = 0.0;
            _swDrips = 0.0;
            _hwDrips = 0.0;
            _TotalBatt = 0.0;
            _Batt1 = 0.0;
            _Batt2 = 0.0;
            _ActiveEnergyDrainRate = 0.0;
            _AdminReport = false;
            _NetType = string.Empty;
    }


        public string NetType
        {
            get { return _NetType; }
            set { _NetType = value; }
        }

        public string GetDescription(Win10OSType input) 
        {
            Type type = input.GetType();
            MemberInfo[] memInfo = type.GetMember(input.ToString());

            if (memInfo != null && memInfo.Length > 0)
            {
                object[] attrs = (object[])memInfo[0].GetCustomAttributes(typeof(DescriptionAttribute), false);
                if (attrs != null && attrs.Length > 0)
                {
                    return ((DescriptionAttribute)attrs[0]).Description;
                }
            }

            return input.ToString();
        }



        public DateTime Date
        {
            get { return _date; }
            set { _date = value; }
        }

        public string DeviceName
        {
            get { return _deviceName; }
            set { _deviceName = value; }
        }

        public string RelativePath
        {
            get { return _relativePath; }
            set { _relativePath = value; }
        }

        public string OsType
        {
            get { return _osType; }
            set { _osType = value; }
        }

        public PowerSliderType PowerSliderMode
        {
            get { return _powerSliderMode; }
            set { _powerSliderMode = value; }
        }

        public PowerSourceType PowerSourceType
        {
            get { return _powerSourceType; }
            set { _powerSourceType = value; }
        }

        public bool HasMemoryDump
        {
            get { return _hasMemoryDump; }
            set { _hasMemoryDump = value; }
        }   
        public bool AdminReport
        {
            get { return _AdminReport; }
            set { _AdminReport = value; }
        }

        public string MemoryDumpFlag
        {
            get { return _memoryDumpFlag; }
            set { _memoryDumpFlag = value; }
        }

        public List<string> BugCheckCode
        {
            get { return _bugCheckCode; }
            set { _bugCheckCode = value; }
        }

        public List<string> MemoryDumpList
        {
            get { return _memoryDumpList; }
            set { _memoryDumpList = value; }
        }

        public bool HasLiveKernelDump
        {
            get { return _hasLiveKernelDump; }
            set { _hasLiveKernelDump = value; }
        }

        public string LiveKernelDumpFlag
        {
            get { return _liveKernelDumpFlag; }
            set { _liveKernelDumpFlag = value; }
        }

        public List<string> LiveKernelCode
        {
            get { return _liveKernelCode; }
            set { _liveKernelCode = value; }
        }

        public List<string> LiveKernelList
        {
            get { return _liveKernelList; }
            set { _liveKernelList = value; }
        }

        public double EnergyDrainRate
        {
            get { return _energyDrainRate; }
            set { _energyDrainRate = value; }
        }

        public double C2ActivePercent
        {
            get { return _c2ActivePercent; }
            set { _c2ActivePercent = value; }
        }

        public double C3ActivePercent
        {
            get { return _c3ActivePercent; }
            set { _c3ActivePercent = value; }
        }

        public double C6ActivePercent
        {
            get { return _c6ActivePercent; }
            set { _c6ActivePercent = value; }
        }

        public double C7ActivePercent
        {
            get { return _c7ActivePercent; }
            set { _c7ActivePercent = value; }
        }

        public double C8ActivePercent
        {
            get { return _c8ActivePercent; }
            set { _c8ActivePercent = value; }
        }

        public double C9ActivePercent
        {
            get { return _c9ActivePercent; }
            set { _c9ActivePercent = value; }
        }

        public double C10ActivePercent
        {
            get { return _c10ActivePercent; }
            set { _c10ActivePercent = value; }
        }

        public double SleepS0Percent
        {
            get { return _sleepS0Percent; }
            set { _sleepS0Percent = value; }
        }

        public double hwDrips
        {
            get { return _hwDrips; }
            set { _hwDrips = value; }
        }
        public double swDrips
        {
            get { return _swDrips; }
            set { _swDrips = value; }
        }     
        public double TotalBatt
        {
            get { return _TotalBatt; }
            set { _TotalBatt = value; }
        }
        public double Batt1
        {
            get { return _Batt1; }
            set { _Batt1 = value; }
        }
        public double Batt2
        {
            get { return _Batt2; }
            set { _Batt2 = value; }
        }
        public double ActiveEnergyDrainRate
        {
            get { return _ActiveEnergyDrainRate; }
            set { _ActiveEnergyDrainRate = value; }
        }
    }
}
