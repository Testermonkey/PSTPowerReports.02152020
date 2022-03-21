using System;
using System.Linq;
using System.IO;

namespace PSTPowerReports
{
    class Util
    {
        private static object writeLogLock = new object();
        private static string traceFileName = string.Empty;

        public static void Initialize()
        {
            if (!CreateTraceFile(Directory.GetCurrentDirectory(), ref traceFileName))
            {
                Console.WriteLine("Error: Couldn't initialize and create a trace file");
            }
        }

        public static void Trace(string text, params object[] args)
        {
            Trace(string.Format(text, args), true);
        }

        private static void Trace(string message, bool displayOnConsole = true)
        {
            string msg = string.Format("{0}: {1}", DateTime.Now, message);

            if (displayOnConsole)
            {
                Console.WriteLine(msg);
            }

            if (!string.IsNullOrEmpty(traceFileName))
            {
                WriteLog(msg, traceFileName);
            }
        }

        private static void WriteLog(string str, string fName)
        {
            try
            {
                if (!string.IsNullOrEmpty(fName))
                { // do not use trace here
                    if (!File.Exists(fName))
                    {
                        StreamWriter sw;
                        sw = new StreamWriter(fName);
                        sw.Close();
                    }

                    lock (writeLogLock)
                    {
                        using (StreamWriter w = File.AppendText(fName))
                        {
                            w.WriteLine(str);
                        }
                    }
                }
            }
            catch (Exception e)
            {  // do not use trace here
                Console.WriteLine("Exception - WriteLog : {0}", e.Message);
            }
        }

        private static bool CreateTraceFile(string testDirPath, ref string logFilePath)
        {
            try
            {
                if (!string.IsNullOrEmpty(logFilePath))
                {
                    return true;
                }

                if (!string.IsNullOrEmpty(testDirPath))
                {
                    string logFolder = testDirPath + "\\Trace";
                    if (!Directory.Exists(logFolder))
                    {
                        Directory.CreateDirectory(logFolder);
                    }

                    logFilePath = logFolder + "\\PstPwrTrace.txt";
                    if (!string.IsNullOrEmpty(logFilePath))
                    {
                        if (!File.Exists(logFilePath))
                        {
                            WriteLog("", logFilePath);
                        }
                        return true;
                    }
                }
            }
            catch (Exception e)
            {
                Console.WriteLine("{0}: Exception - CreateTraceFile: {1}", DateTime.Now, e.Message);
            }
            return false;
        }
    }
}
