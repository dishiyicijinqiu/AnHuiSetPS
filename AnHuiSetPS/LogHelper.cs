using Microsoft.Practices.EnterpriseLibrary.Logging;
using System;
using System.Diagnostics;

namespace AnHuiSetPS
{
    public class LogHelper
    {
        public static void WriteError(string Category, string ErrorMsg)
        {
            LogEntry entry = new LogEntry();
            entry.Message = ErrorMsg;
            entry.TimeStamp = DateTime.Now;
            entry.Categories = new string[] { Category };
            Logger.Write(entry);
        }
    }
}
