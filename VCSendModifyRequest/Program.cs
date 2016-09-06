using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Net.Http;
using System.Net;
using System.Collections.Specialized;
using System.IO;
using Excel = Microsoft.Office.Interop.Excel;

namespace VCSendModifyRequest
{
   public class Program
   {
      public class RequestModel
      {
         public string boardId { get; set; }             = String.Empty;
         public string mode { get; set; }                = String.Empty;
         public string bdId { get; set; }                = String.Empty;
         public string bdRank { get; set; }              = String.Empty;
         public string bdActive { get; set; }            = String.Empty;
         public string bdPremium { get; set; }           = String.Empty;
         public string bdPremiumStart { get; set; }      = String.Empty;
         public string bdPremiumEnd { get; set; }        = String.Empty;
         public string bdMainDisplay { get; set; }       = String.Empty;
         public string old_imageFile1 { get; set; }      = String.Empty;
         public string old_imageFile2 { get; set; }      = String.Empty;
         public string old_imageFile3 { get; set; }      = String.Empty;
         public string old_imageFile4 { get; set; }      = String.Empty;
         public string old_imageFile5 { get; set; }      = String.Empty;
         public string old_imageFile6 { get; set; }      = String.Empty;
         public string old_imageFile7 { get; set; }      = String.Empty;
         public string old_imageFile8 { get; set; }      = String.Empty;
         public string old_imageFile9 { get; set; }      = String.Empty;
         public string old_imageFile10 { get; set; }     = String.Empty;
         public string bdIpAddress { get; set; }         = String.Empty;
         public string bdTitle { get; set; }             = String.Empty;
         public string bdName { get; set; }              = String.Empty;
         public string bdPassword { get; set; }          = String.Empty;
         public string bdEmail { get; set; }             = String.Empty;
         public string bdPhone { get; set; }             = String.Empty;
         public string bdType { get; set; }              = String.Empty;
         public string bdLocation { get; set; }          = String.Empty;
         public string bdPrice { get; set; }             = String.Empty;
         public string bdTag { get; set; }               = String.Empty;
         public string part { get; set; }                = String.Empty;
         public string bdSummary { get; set; }           = String.Empty;
         public string bdLink { get; set; }              = String.Empty;
         public string chk_image { get; set; }           = String.Empty;
         public string photo1 { get; set; }              = String.Empty;
         public string photo2 { get; set; }              = String.Empty;
         public string photo3 { get; set; }              = String.Empty;
         public string photo4 { get; set; }              = String.Empty;
         public string photo5 { get; set; }              = String.Empty;
         public string photo6 { get; set; }              = String.Empty;
         public string photo7 { get; set; }              = String.Empty;
         public string photo8 { get; set; }              = String.Empty;
         public string photo9 { get; set; }              = String.Empty;
         public string photo10 { get; set; }             = String.Empty;
         public string bdDescription { get; set; } = String.Empty;


      }

      static void Main()
      {
         const string url = @"http://www.vanchosun.com/market/m_tutor/Function_tutor.php";

         //Check if the file exists
         string fileName = "Data.xls";
         CheckDataFile(fileName);

         //Read File here
         List<RequestModel> reqList = ReadDataFile(fileName);

         Console.WriteLine("Reading inputs..");
         try
         {
            foreach (RequestModel model in reqList)
            {
               SendPostRequest(url, model.boardId, model.mode, model.bdId, model.bdRank, model.bdActive, model.bdPremium, model.bdPremiumStart,
                               model.bdPremiumEnd, model.bdMainDisplay, model.old_imageFile1, model.old_imageFile2, model.old_imageFile3, model.old_imageFile4,
                               model.old_imageFile5, model.old_imageFile6, model.old_imageFile7, model.old_imageFile8, model.old_imageFile9, model.old_imageFile10,
                               model.bdIpAddress, model.bdTitle, model.bdName, model.bdPassword, model.bdEmail, model.bdPhone, model.bdType, model.bdLocation,
                               model.bdPrice, model.bdTag, model.part, model.bdSummary, model.bdLink, model.chk_image, model.photo1, model.photo2, model.photo3,
                                model.photo4, model.photo5, model.photo6, model.photo7, model.photo8, model.photo9, model.photo10, model.bdDescription);
                             
              //SendPostRequest(url,  "6", "update_tutor", "25757", "0", "1", "0", "0000-00-00 00:00:00", "0000-00-00 00:00:00", "0",
              //  "tutor1_1431466420.jpg", "tutor2_1431466421.jpg", "", "", "", "", "", "", "", "", "201.129.31.55", "=●●■포토샵■일러스트레이터■인디자인■디지털페인팅■●●=", "ARTSTUDIOHELP", "popopo", "artstudiohelp@gmail.com", "778-847-1700", "학원", "5", "", "", "7", "", "", "1", "", "", "", "", "", "", "", "", "", "", desc);
               System.Threading.Thread.Sleep(1000);
            }

         }
         catch
         {
            Console.WriteLine("Invalid input..");
         }
         Console.WriteLine("Exiting..");
      }

      private static void CheckDataFile(string fileName)
      {
         if (!System.IO.File.Exists(fileName))
         {
            Console.WriteLine("There is no " + fileName + " in the directory.");
            return;
         }
      }

      private static List<RequestModel> ReadDataFile(string fileName)
      {
         Excel.Application xlApp;
         Excel.Workbook xlWorkBook;
         Excel.Worksheet xlWorkSheet;
         Excel.Range range;

         List<RequestModel> reqList = new List<RequestModel>();

         try
         {

            xlApp = new Excel.Application();
            xlWorkBook = xlApp.Workbooks.Open(AppDomain.CurrentDomain.BaseDirectory + fileName, 0, true, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
            xlWorkSheet = xlWorkBook.Worksheets["Data1"];

            range = xlWorkSheet.UsedRange;

            int colNo = xlWorkSheet.UsedRange.Columns.Count;
            int rowNo = xlWorkSheet.UsedRange.Rows.Count;

            object[,] array = xlWorkSheet.UsedRange.Value;

            for (int j = 2; j <= colNo; j++)
            {
               RequestModel model = new RequestModel();
               for (int i = 1; i <= rowNo; i++)
               {
                  if (array[i, j] != null)
                  {
                     if (array[i, 1].ToString() == "boardId")
                        model.boardId = array[i, j].ToString();
                     if (array[i, 1].ToString() == "mode")
                        model.mode = array[i, j].ToString();
                     if (array[i, 1].ToString() == "bdId")
                        model.bdId = array[i, j].ToString();
                     if (array[i, 1].ToString() == "bdRank")
                        model.bdRank = array[i, j].ToString();
                     if (array[i, 1].ToString() == "bdActive")
                        model.bdActive = array[i, j].ToString();
                     if (array[i, 1].ToString() == "bdPremium")
                        model.bdPremium = array[i, j].ToString();
                     if (array[i, 1].ToString() == "bdPremiumStart")
                        model.bdPremiumStart = array[i, j].ToString();
                     if (array[i, 1].ToString() == "bdPremiumEnd")
                        model.bdPremiumEnd = array[i, j].ToString();
                     if (array[i, 1].ToString() == "bdMainDisplay")
                        model.bdMainDisplay = array[i, j].ToString();
                     if (array[i, 1].ToString() == "old_imageFile1")
                        model.old_imageFile1 = array[i, j].ToString();
                     if (array[i, 1].ToString() == "old_imageFile2")
                        model.old_imageFile2 = array[i, j].ToString();
                     if (array[i, 1].ToString() == "old_imageFile3")
                        model.old_imageFile3 = array[i, j].ToString();
                     if (array[i, 1].ToString() == "old_imageFile4")
                        model.old_imageFile4 = array[i, j].ToString();
                     if (array[i, 1].ToString() == "old_imageFile5")
                        model.old_imageFile5 = array[i, j].ToString();
                     if (array[i, 1].ToString() == "old_imageFile6")
                        model.old_imageFile6 = array[i, j].ToString();
                     if (array[i, 1].ToString() == "old_imageFile7")
                        model.old_imageFile7 = array[i, j].ToString();
                     if (array[i, 1].ToString() == "old_imageFile8")
                        model.old_imageFile8 = array[i, j].ToString();
                     if (array[i, 1].ToString() == "old_imageFile9")
                        model.old_imageFile9 = array[i, j].ToString();
                     if (array[i, 1].ToString() == "old_imageFile10")
                        model.old_imageFile10 = array[i, j].ToString();
                     if (array[i, 1].ToString() == "bdIpAddress")
                        model.bdIpAddress = array[i, j].ToString();
                     if (array[i, 1].ToString() == "bdTitle")
                        model.bdTitle = array[i, j].ToString();
                     if (array[i, 1].ToString() == "bdName")
                        model.bdName = array[i, j].ToString();
                     if (array[i, 1].ToString() == "bdPassword")
                        model.bdPassword = array[i, j].ToString();
                     if (array[i, 1].ToString() == "bdEmail")
                        model.bdEmail = array[i, j].ToString();
                     if (array[i, 1].ToString() == "bdPhone")
                        model.bdPhone = array[i, j].ToString();
                     if (array[i, 1].ToString() == "bdType")
                        model.bdType = array[i, j].ToString();
                     if (array[i, 1].ToString() == "bdLocation")
                        model.bdLocation = array[i, j].ToString();
                     if (array[i, 1].ToString() == "bdPrice")
                        model.bdPrice = array[i, j].ToString();
                     if (array[i, 1].ToString() == "bdTag")
                        model.bdTag = array[i, j].ToString();
                     if (array[i, 1].ToString() == "part")
                        model.part = array[i, j].ToString();
                     if (array[i, 1].ToString() == "bdSummary")
                        model.bdSummary = array[i, j].ToString();
                     if (array[i, 1].ToString() == "bdLink")
                        model.bdLink = array[i, j].ToString();
                     if (array[i, 1].ToString() == "chk_image")
                        model.chk_image = array[i, j].ToString();
                     if (array[i, 1].ToString() == "photo1")
                        model.photo1 = array[i, j].ToString();
                     if (array[i, 1].ToString() == "photo2")
                        model.photo2 = array[i, j].ToString();
                     if (array[i, 1].ToString() == "photo3")
                        model.photo3 = array[i, j].ToString();
                     if (array[i, 1].ToString() == "photo4")
                        model.photo4 = array[i, j].ToString();
                     if (array[i, 1].ToString() == "photo5")
                        model.photo5 = array[i, j].ToString();
                     if (array[i, 1].ToString() == "photo6")
                        model.photo6 = array[i, j].ToString();
                     if (array[i, 1].ToString() == "photo7")
                        model.photo7 = array[i, j].ToString();
                     if (array[i, 1].ToString() == "photo8")
                        model.photo8 = array[i, j].ToString();
                     if (array[i, 1].ToString() == "photo9")
                        model.photo9 = array[i, j].ToString();
                     if (array[i, 1].ToString() == "photo10")
                        model.photo10 = array[i, j].ToString();
                     if (array[i, 1].ToString() == "bdDescription")
                        model.bdDescription = array[i, j].ToString();
                  }
               }
               reqList.Add(model);
            }

            xlWorkBook.Close(true, null, null);
            xlApp.Quit();

            releaseObject(xlWorkSheet);
            releaseObject(xlWorkBook);
            releaseObject(xlApp);
         }
         catch
         {

         }
 
         return reqList;
      }

      static private void releaseObject(object obj)
      {
         try
         {
            System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
            obj = null;
         }
         catch (Exception ex)
         {
            obj = null;
            Console.WriteLine("Unable to release the Object " + ex.ToString());
         }
         finally
         {
            GC.Collect();
         }
      }

      private static string SendPostRequest( string url,
         string boardId, string mode, string bdId, string bdRank, string bdActive,
         string bdPremium, string bdPremiumStart, string bdPremiumEnd, string bdMainDisplay,
         string old_imageFile1, string old_imageFile2, string old_imageFile3, string old_imageFile4, string old_imageFile5,
         string old_imageFile6, string old_imageFile7, string old_imageFile8, string old_imageFile9, string old_imageFile10,
         string bdIpAddress, string bdTitle, string bdName, string bdPassword, string bdEmail,
         string bdPhone, string bdType, string bdLocation, string bdPrice, string bdTag,
         string part, string bdSummary, string bdLink, string chk_image, string photo1,
         string photo2, string photo3, string photo4, string photo5, string photo6,
         string photo7, string photo8, string photo9, string photo10, string bdDescription)
      {
         var chars = "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789";
         var requestFormat = "Content-Disposition: form-data; name=\"{0}\"\r\n\r\n{1}";
         var requestFormatWithFile = "Content-Disposition: form-data; name=\"{0}\"; filename=\"{1}\";"
                        + "\r\nContent-Type: application/octet-stream\r\n\r\n{2}";

         var random = new Random();
         var boundary = "----WebKitFormBoundary" + new string(
             Enumerable.Repeat(chars, 16)
                       .Select(s => s[random.Next(s.Length)])
                       .ToArray());


         //byte[] bytes = Encoding.UTF8.GetBytes(querystring);
         HttpWebRequest request = (HttpWebRequest)WebRequest.Create(url);
         request.ContentType = "multipart/form-data; boundary=" + boundary;
         request.Method = "POST";
         request.Accept = "text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8";
         request.KeepAlive = true;
         request.Referer = "http://www.vanchosun.com/market/main/frame.php";
         
         StreamWriter requestWriter = new StreamWriter(request.GetRequestStream());

         try
         {
            requestWriter.Write("\r\n--" + boundary + "\r\n");
            requestWriter.Write(requestFormat, "boardId", boardId);
            requestWriter.Write("\r\n--" + boundary + "\r\n");
            requestWriter.Write(requestFormat, "mode", mode);
            requestWriter.Write("\r\n--" + boundary + "\r\n");
            requestWriter.Write(requestFormat, "bdId", bdId);
            requestWriter.Write("\r\n--" + boundary + "\r\n");
            requestWriter.Write(requestFormat, "bdRank", bdRank);
            requestWriter.Write("\r\n--" + boundary + "\r\n");
            requestWriter.Write(requestFormat, "bdActive", bdActive);
            requestWriter.Write("\r\n--" + boundary + "\r\n");
            requestWriter.Write(requestFormat, "bdPremium", bdPremium);
            requestWriter.Write("\r\n--" + boundary + "\r\n");
            requestWriter.Write(requestFormat, "bdPremiumStart", bdPremiumStart);
            requestWriter.Write("\r\n--" + boundary + "\r\n");
            requestWriter.Write(requestFormat, "bdPremiumEnd", bdPremiumEnd);
            requestWriter.Write("\r\n--" + boundary + "\r\n");
            requestWriter.Write(requestFormat, "bdMainDisplay", bdMainDisplay);
            requestWriter.Write("\r\n--" + boundary + "\r\n");
            requestWriter.Write(requestFormat, "old_imageFile1", old_imageFile1);
            requestWriter.Write("\r\n--" + boundary + "\r\n");
            requestWriter.Write(requestFormat, "old_imageFile2", old_imageFile2);
            requestWriter.Write("\r\n--" + boundary + "\r\n");
            requestWriter.Write(requestFormat, "old_imageFile3", old_imageFile3);
            requestWriter.Write("\r\n--" + boundary + "\r\n");
            requestWriter.Write(requestFormat, "old_imageFile4", old_imageFile4);
            requestWriter.Write("\r\n--" + boundary + "\r\n");
            requestWriter.Write(requestFormat, "old_imageFile5", old_imageFile5);
            requestWriter.Write("\r\n--" + boundary + "\r\n");
            requestWriter.Write(requestFormat, "old_imageFile6", old_imageFile6);
            requestWriter.Write("\r\n--" + boundary + "\r\n");
            requestWriter.Write(requestFormat, "old_imageFile7", old_imageFile7);
            requestWriter.Write("\r\n--" + boundary + "\r\n");
            requestWriter.Write(requestFormat, "old_imageFile8", old_imageFile8);
            requestWriter.Write("\r\n--" + boundary + "\r\n");
            requestWriter.Write(requestFormat, "old_imageFile9", old_imageFile9);
            requestWriter.Write("\r\n--" + boundary + "\r\n");
            requestWriter.Write(requestFormat, "old_imageFile10", old_imageFile10);
            requestWriter.Write("\r\n--" + boundary + "\r\n");
            requestWriter.Write(requestFormat, "bdIpAddress", bdIpAddress);
            requestWriter.Write("\r\n--" + boundary + "\r\n");
            requestWriter.Write(requestFormat, "bdTitle", bdTitle);
            requestWriter.Write("\r\n--" + boundary + "\r\n");
            requestWriter.Write(requestFormat, "bdName", bdName);
            requestWriter.Write("\r\n--" + boundary + "\r\n");
            requestWriter.Write(requestFormat, "bdPassword", bdPassword);
            requestWriter.Write("\r\n--" + boundary + "\r\n");
            requestWriter.Write(requestFormat, "bdEmail", bdEmail);
            requestWriter.Write("\r\n--" + boundary + "\r\n");
            requestWriter.Write(requestFormat, "bdPhone", bdPhone);
            requestWriter.Write("\r\n--" + boundary + "\r\n");
            requestWriter.Write(requestFormat, "bdType", bdType);
            requestWriter.Write("\r\n--" + boundary + "\r\n");
            requestWriter.Write(requestFormat, "bdLocation", bdLocation);
            requestWriter.Write("\r\n--" + boundary + "\r\n");
            requestWriter.Write(requestFormat, "bdPrice", bdPrice);
            requestWriter.Write("\r\n--" + boundary + "\r\n");
            requestWriter.Write(requestFormat, "bdTag", bdTag);
            requestWriter.Write("\r\n--" + boundary + "\r\n");
            requestWriter.Write(requestFormat, "part", part);
            requestWriter.Write("\r\n--" + boundary + "\r\n");
            requestWriter.Write(requestFormat, "bdSummary", bdSummary);
            requestWriter.Write("\r\n--" + boundary + "\r\n");
            requestWriter.Write(requestFormat, "bdLink", bdLink);
            requestWriter.Write("\r\n--" + boundary + "\r\n");
            requestWriter.Write(requestFormat, "chk_image", chk_image);
            requestWriter.Write("\r\n--" + boundary + "\r\n");
            requestWriter.Write(requestFormatWithFile, "photo1", photo1, "");
            requestWriter.Write("\r\n--" + boundary + "\r\n");
            requestWriter.Write(requestFormatWithFile, "photo2", photo2, "");
            requestWriter.Write("\r\n--" + boundary + "\r\n");
            requestWriter.Write(requestFormatWithFile, "photo3", photo3, "");
            requestWriter.Write("\r\n--" + boundary + "\r\n");
            requestWriter.Write(requestFormatWithFile, "photo4", photo4, "");
            requestWriter.Write("\r\n--" + boundary + "\r\n");
            requestWriter.Write(requestFormatWithFile, "photo5", photo5, "");
            requestWriter.Write("\r\n--" + boundary + "\r\n");
            requestWriter.Write(requestFormatWithFile, "photo6", photo6, "");
            requestWriter.Write("\r\n--" + boundary + "\r\n");
            requestWriter.Write(requestFormatWithFile, "photo7", photo7, "");
            requestWriter.Write("\r\n--" + boundary + "\r\n");
            requestWriter.Write(requestFormatWithFile, "photo8", photo8, "");
            requestWriter.Write("\r\n--" + boundary + "\r\n");
            requestWriter.Write(requestFormatWithFile, "photo9", photo9, "");
            requestWriter.Write("\r\n--" + boundary + "\r\n");
            requestWriter.Write(requestFormatWithFile, "photo10", photo10, "");
            requestWriter.Write("\r\n--" + boundary + "\r\n");
            requestWriter.Write(requestFormat, "bdDescription", bdDescription);
            requestWriter.Write("\r\n--" + boundary + "\r\n");
         }
         catch
         {
            throw;
         }
         finally
         {
            requestWriter.Close();
            requestWriter = null;
         }

         HttpWebResponse res = (HttpWebResponse)request.GetResponse();
         using (StreamReader reader = new StreamReader(res.GetResponseStream()))
         {
            return reader.ReadToEnd();
         }
      }

   }
}
