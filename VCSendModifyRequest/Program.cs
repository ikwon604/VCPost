using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Net.Http;
using System.Net;
using System.Collections.Specialized;
using System.IO;
using VCSendModifyRequest.Model;

namespace VCSendModifyRequest
{
   public class Program
   {
      static void Main()
      {
         //Check if the file exists
         RequestModel model = new RequestModel();
         string s = model.ReadTargetPostRequest();
         //Read File here
         Console.WriteLine("Reading inputs..");
         try
         {

         }
         catch
         {
            Console.WriteLine("Invalid input..");
         }
         Console.WriteLine("Exiting..");
      }

      

   }
}
