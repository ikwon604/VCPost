using System;
using VCSendModifyRequest.Model;

namespace VCSendModifyRequest
{
   public class Program
   {
      static void Main()
      {
         try
         {
            TraceLog.Instance.WriteLine("Starting the program...");
            RequestUtil util = new RequestUtil();
            util.UpdatePosts();            
         }
         catch (Exception e){
            TraceLog.Instance.WriteLine(string.Format("Error: {0}", e.Message));
         }
         TraceLog.Instance.WriteLine("Exiting...");
      }

      

   }
}
