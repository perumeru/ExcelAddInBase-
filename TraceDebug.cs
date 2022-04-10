using System.Diagnostics;
using System.IO;

namespace ExcelAddIn1
{
    internal class TraceDebug
    {
        [Conditional("ELOG")]
        public static void Create()
        {
            FileStream fileStream = new FileStream(Directory.GetCurrentDirectory() + @"\\ErrorLog.txt", FileMode.Append);
            StreamWriter sw = new StreamWriter(fileStream);
            sw.AutoFlush = true;
            TextWriter tw = TextWriter.Synchronized(sw);
            TextWriterTraceListener twtl = new TextWriterTraceListener(tw, "LogFile");
            Trace.Listeners.Add(twtl);
        }
        [Conditional("ELOG")]
        public static void Dispose()
        {
            Trace.Close();
        }
    }
}
