using DocumentFormat.OpenXml.Office2010.ExcelAc;
using System;
using System.Collections.Generic;

namespace PSTPowerReports
{
    class Program
    {
        static PSTPowerLogs pstLogs = new PSTPowerLogs();

        static void Main(string[] args)
        {
#if DEBUG
            List<string> testargs = new List<string>();
            testargs.Add("/h1=19H2_12.test.0");
            testargs.Add("/h2=Test");
            //testargs.Add(@"/dir=\\10.200.249.193\CS_Stress\StressTesting\Cruz-OEMCA\JanTP2021\19H2\12.343.0\Jan21-HCL");
            testargs.Add(@"/dir=\\10.200.249.193\CS_Stress\StressTesting\Cayucos-OEMBY\Jun2021-SVTP2\21H2\6.510.0");
            //testargs.Add("false");
            args = testargs.ToArray();

#endif
            try
            {
                if (pstLogs.Initialize(args) == 1)
                {
                    if (!pstLogs.Process(out bool foundException))
                    {
                        if (foundException)
                        {
                            Util.Trace("Error: Something went wrong.... Could not process PST logs");
                        }
                        else
                        {
                            Util.Trace("Warning: Could not find PST test results folder");
                        }
                    }
                    else
                    {
                        if (!pstLogs.GenerateReport(out bool foundReportException))
                        {
                            if (foundReportException)
                            {
                                Util.Trace("Error: Something went wrong.... Could not generate PST report");
                            }
                        }
                    }
                }
            }
            catch(Exception e)
            {
                Util.Trace("Exception in Main -> {0}", e.Message);
            }
        }
    }
}
