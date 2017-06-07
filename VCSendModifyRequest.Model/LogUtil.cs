using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace VCSendModifyRequest.Model
{
   public class TraceLog
   {
      private const string TraceLogFile = "VCRequest_Trace.log";
      private const string TraceCategory = "VCRequest_RunTime";

      #region Singleton Pattern
      //
      // Using version 6 in the following singleton design pattern:
      // http://csharpindepth.com/Articles/General/Singleton.aspx
      // Advantage is it's thread-safe and lazy initialization
      //
      private static readonly Lazy<TraceLog> m_instance = new Lazy<TraceLog>(() => new TraceLog());

      /// <summary>
      /// Get the singleton instance
      /// </summary>
      public static TraceLog Instance
      {
         get { return m_instance.Value; }
      }

      /// <summary>
      /// Constructor
      /// </summary>
      private TraceLog()
      {
         Init();
      }
      #endregion

      /// <summary>
      /// Write a new line of message
      /// </summary>
      /// <param name="message">the message to write</param>
      /// <param name="time">whether time is included in the message</param>
      /// <param name="category">the category this message belongs to</param>
      public void WriteLine(string message, bool time = true, string category = TraceCategory)
      {
         Console.WriteLine(time ? TimeFormat + " " + message : message, category);
         Trace.WriteLine(time ? TimeFormat + " " + message : message, category);
      }

      private void Init()
      {
         // Init Trace
         Trace.Listeners.Add(new TextWriterTraceListener(TraceLogFile));
         Trace.AutoFlush = true;
      }

      private string TimeFormat
      {
         get { return "[" + DateTime.Now.ToString() + "]"; }
      }
   }
}
