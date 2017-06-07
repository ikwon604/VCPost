using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using HtmlAgilityPack;


namespace VCSendModifyRequest.Model
{
   public class BoardIDModel
   {
      public string boardId { get; set; } = String.Empty;
      public string bdId { get; set; } = String.Empty;
      public string bdPassword { get; set; } = String.Empty;
   }

   public class PostRequestModel : BoardIDModel
   {
      public string tmpPremium { get; set; } = String.Empty;
      public string mode { get; set; } = String.Empty;
      public string bdRank { get; set; } = String.Empty;
      public string bdActive { get; set; } = String.Empty;
      public string bdPremium { get; set; } = String.Empty;
      public string bdPremiumStart { get; set; } = String.Empty;
      public string bdPremiumEnd { get; set; } = String.Empty;
      public string bdMainDisplay { get; set; } = String.Empty;
      public string old_imageFile1 { get; set; } = String.Empty;
      public string old_imageFile2 { get; set; } = String.Empty;
      public string old_imageFile3 { get; set; } = String.Empty;
      public string old_imageFile4 { get; set; } = String.Empty;
      public string old_imageFile5 { get; set; } = String.Empty;
      public string old_imageFile6 { get; set; } = String.Empty;
      public string old_imageFile7 { get; set; } = String.Empty;
      public string old_imageFile8 { get; set; } = String.Empty;
      public string old_imageFile9 { get; set; } = String.Empty;
      public string old_imageFile10 { get; set; } = String.Empty;
      public string bdIpAddress { get; set; } = String.Empty;
      public string bdTitle { get; set; } = String.Empty;
      public string bdName { get; set; } = String.Empty;
      public string bdEmail { get; set; } = String.Empty;
      public string bdPhone { get; set; } = String.Empty;
      public string bdType { get; set; } = String.Empty;
      public string bdLocation { get; set; } = String.Empty;
      public string bdPrice { get; set; } = String.Empty;
      public string bdTag { get; set; } = String.Empty;
      public string chk_image { get; set; } = String.Empty;
      public string photo1 { get; set; } = String.Empty;
      public string photo2 { get; set; } = String.Empty;
      public string photo3 { get; set; } = String.Empty;
      public string photo4 { get; set; } = String.Empty;
      public string photo5 { get; set; } = String.Empty;
      public string photo6 { get; set; } = String.Empty;
      public string photo7 { get; set; } = String.Empty;
      public string photo8 { get; set; } = String.Empty;
      public string photo9 { get; set; } = String.Empty;
      public string photo10 { get; set; } = String.Empty;
      public string bdDescription { get; set; } = String.Empty;
   }

   public class RequestUtil
   {
      const string targetUrl = @"http://www.vanchosun.com/market/main/frame.php";
      const string updateUrl = @"http://www.vanchosun.com/market/m_tutor/Function_tutor.php";
      const string filename = @"Data.xls";

      public void UpdatePosts()
      {
         List<BoardIDModel> list = new List<BoardIDModel>();
         TraceLog.Instance.WriteLine(string.Format("Start reading a data file: {0}", filename));
         list = ReadDataFile(filename);
         foreach (BoardIDModel model in list)
         {
            TraceLog.Instance.WriteLine(string.Format("Getting info from the target post {0}", model.bdId));
            PostRequestModel postModel = ReadTargetPostRequest(model.boardId, model.bdId, model.bdPassword);
            TraceLog.Instance.WriteLine(string.Format("Updating the target post {0}", model.bdId));
            SendPostRequest(postModel);
            TraceLog.Instance.WriteLine(string.Format("Done updating the target post {0}", model.bdId));
         }
      }

      private PostRequestModel ReadTargetPostRequest(string boardId, string bdId, string bdPassword)
      {
         HttpWebRequest request = (HttpWebRequest)WebRequest.Create(targetUrl);
         request.Method = "POST";
         request.Host = "www.vanchosun.com";
         request.KeepAlive = true;
         request.ContentType = @"application/x-www-form-urlencoded";
         request.Accept = "text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8";
         request.Referer = "http://www.vanchosun.com/market/main/frame.php?main=tutor&bdId=28628&cpage1=1&search_type=&search_title=&search_location=";
         try
         {
            var htmlDoc = new HtmlDocument();
            using (StreamWriter requestWriter = new StreamWriter(request.GetRequestStream()))
            {
               requestWriter.Write(@"main=tutor&sub=tutor%28write%29&boardId={0}&bdId={1}&cpage1=1&bdPassword={2}", boardId, bdId, bdPassword);
            }
            HttpWebResponse res = (HttpWebResponse)request.GetResponse();
            using (StreamReader reader = new StreamReader(res.GetResponseStream()))
            {
               htmlDoc.LoadHtml(reader.ReadToEnd());
               var htmlNodes = htmlDoc.DocumentNode.SelectNodes("//div");
               PostRequestModel model = new PostRequestModel();

               HtmlNode targetNode = htmlNodes.Where(x => x.Id == "cf_middle").First();
               List<HtmlNode> targetList = targetNode.Descendants("input").ToList();

               model.boardId = targetList.Where(x => x.Attributes["name"].Value == "boardId").First().Attributes["value"].Value;
               model.tmpPremium = targetList.Where(x => x.Attributes["name"].Value == "tmpPremium").First().Attributes["value"].Value;
               //hard code this
               model.mode = "update_tutor";
               model.bdId = targetList.Where(x => x.Attributes["name"].Value == "bdId").First().Attributes["value"].Value;
               model.bdRank = targetList.Where(x => x.Attributes["name"].Value == "bdRank").First().Attributes["value"].Value;
               model.bdActive = targetList.Where(x => x.Attributes["name"].Value == "bdActive").First().Attributes["value"].Value;
               model.bdPremium = targetList.Where(x => x.Attributes["name"].Value == "bdPremium").First().Attributes["value"].Value;
               model.bdPremiumStart = targetList.Where(x => x.Attributes["name"].Value == "bdPremiumStart").First().Attributes["value"].Value;
               model.bdPremiumEnd = targetList.Where(x => x.Attributes["name"].Value == "bdPremiumEnd").First().Attributes["value"].Value;
               model.bdMainDisplay = targetList.Where(x => x.Attributes["name"].Value == "bdMainDisplay").First().Attributes["value"].Value;
               model.old_imageFile1 = targetList.Where(x => x.Attributes["name"].Value == "old_imageFile1").First().Attributes["value"].Value;
               model.old_imageFile2 = targetList.Where(x => x.Attributes["name"].Value == "old_imageFile2").First().Attributes["value"].Value;
               model.old_imageFile3 = targetList.Where(x => x.Attributes["name"].Value == "old_imageFile3").First().Attributes["value"].Value;
               model.old_imageFile4 = targetList.Where(x => x.Attributes["name"].Value == "old_imageFile4").First().Attributes["value"].Value;
               model.old_imageFile5 = targetList.Where(x => x.Attributes["name"].Value == "old_imageFile5").First().Attributes["value"].Value;
               model.old_imageFile6 = targetList.Where(x => x.Attributes["name"].Value == "old_imageFile6").First().Attributes["value"].Value;
               model.old_imageFile7 = targetList.Where(x => x.Attributes["name"].Value == "old_imageFile7").First().Attributes["value"].Value;
               model.old_imageFile8 = targetList.Where(x => x.Attributes["name"].Value == "old_imageFile8").First().Attributes["value"].Value;
               model.old_imageFile9 = targetList.Where(x => x.Attributes["name"].Value == "old_imageFile9").First().Attributes["value"].Value;
               model.old_imageFile10 = targetList.Where(x => x.Attributes["name"].Value == "old_imageFile10").First().Attributes["value"].Value;
               model.bdIpAddress = targetList.Where(x => x.Attributes["name"].Value == "bdIpAddress").First().Attributes["value"].Value;
               model.bdTitle = targetList.Where(x => x.Attributes["name"].Value == "bdTitle").First().Attributes["value"].Value;
               model.bdName = targetList.Where(x => x.Attributes["name"].Value == "bdName").First().Attributes["value"].Value;
               model.bdPassword = targetList.Where(x => x.Attributes["name"].Value == "bdPassword").First().Attributes["value"].Value;
               model.bdPhone = targetList.Where(x => x.Attributes["name"].Value == "bdPhone").First().Attributes["value"].Value;
               model.bdEmail = targetList.Where(x => x.Attributes["name"].Value == "bdEmail").First().Attributes["value"].Value;
               model.bdType = targetNode.Descendants("select").Where(x => x.Attributes["name"].Value == "bdType").First().Descendants("option").First().Attributes["value"].Value;
               model.bdLocation = targetNode.Descendants("select").Where(x => x.Attributes["name"].Value == "bdLocation").First().SelectNodes("option[@selected]").First().Attributes["value"].Value;
               model.bdPrice = targetList.Where(x => x.Attributes["name"].Value == "bdPrice").First().Attributes["value"].Value;
               model.bdTag = targetList.Where(x => x.Attributes["name"].Value == "bdTag").First().Attributes["value"].Value;
               model.chk_image = targetList.Where(x => x.Attributes["name"].Value == "chk_image").First().Attributes["value"].Value;
               model.photo1 = targetList.Where(x => x.Attributes["name"].Value == "photo1").First().Attributes["value"].Value;
               model.photo2 = targetList.Where(x => x.Attributes["name"].Value == "photo2").First().Attributes["value"].Value;
               model.photo3 = targetList.Where(x => x.Attributes["name"].Value == "photo3").First().Attributes["value"].Value;
               model.photo4 = targetList.Where(x => x.Attributes["name"].Value == "photo4").First().Attributes["value"].Value;
               model.photo5 = targetList.Where(x => x.Attributes["name"].Value == "photo5").First().Attributes["value"].Value;
               model.photo6 = targetList.Where(x => x.Attributes["name"].Value == "photo6").First().Attributes["value"].Value;
               model.photo7 = targetList.Where(x => x.Attributes["name"].Value == "photo7").First().Attributes["value"].Value;
               model.photo8 = targetList.Where(x => x.Attributes["name"].Value == "photo8").First().Attributes["value"].Value;
               model.photo9 = targetList.Where(x => x.Attributes["name"].Value == "photo9").First().Attributes["value"].Value;
               model.photo10 = targetList.Where(x => x.Attributes["name"].Value == "photo10").First().Attributes["value"].Value;
               model.bdDescription = targetNode.Descendants("textarea").Where(x => x.Attributes["name"].Value == "bdDescription").First().InnerHtml;

               return model;
            }
         }catch(Exception e){
            TraceLog.Instance.WriteLine(string.Format("Error: {0}", e.Message));
            return null;
         }
      }

      /// <summary>
      /// Update post like following
      /// </summary>
      private string SendPostRequest(PostRequestModel model)
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
         HttpWebRequest request = (HttpWebRequest)WebRequest.Create(updateUrl);
         request.ContentType = "multipart/form-data; boundary=" + boundary;
         request.Method = "POST";
         request.Accept = "text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8";
         request.KeepAlive = true;
         request.Referer = "http://www.vanchosun.com/market/main/frame.php";

         StreamWriter requestWriter = new StreamWriter(request.GetRequestStream());

         try
         {
            requestWriter.Write("\r\n--" + boundary + "\r\n");
            requestWriter.Write(requestFormat, "boardId", model.boardId);
            requestWriter.Write("\r\n--" + boundary + "\r\n");
            requestWriter.Write(requestFormat, "tmpPremium", model.tmpPremium);
            requestWriter.Write("\r\n--" + boundary + "\r\n");
            requestWriter.Write(requestFormat, "mode", model.mode);
            requestWriter.Write("\r\n--" + boundary + "\r\n");
            requestWriter.Write(requestFormat, "bdId", model.bdId);
            requestWriter.Write("\r\n--" + boundary + "\r\n");
            requestWriter.Write(requestFormat, "bdRank", model.bdRank);
            requestWriter.Write("\r\n--" + boundary + "\r\n");
            requestWriter.Write(requestFormat, "bdActive", model.bdActive);
            requestWriter.Write("\r\n--" + boundary + "\r\n");
            requestWriter.Write(requestFormat, "bdPremium", model.bdPremium);
            requestWriter.Write("\r\n--" + boundary + "\r\n");
            requestWriter.Write(requestFormat, "bdPremiumStart", model.bdPremiumStart);
            requestWriter.Write("\r\n--" + boundary + "\r\n");
            requestWriter.Write(requestFormat, "bdPremiumEnd", model.bdPremiumEnd);
            requestWriter.Write("\r\n--" + boundary + "\r\n");
            requestWriter.Write(requestFormat, "bdMainDisplay", model.bdMainDisplay);
            requestWriter.Write("\r\n--" + boundary + "\r\n");
            requestWriter.Write(requestFormat, "old_imageFile1", model.old_imageFile1);
            requestWriter.Write("\r\n--" + boundary + "\r\n");
            requestWriter.Write(requestFormat, "old_imageFile2", model.old_imageFile2);
            requestWriter.Write("\r\n--" + boundary + "\r\n");
            requestWriter.Write(requestFormat, "old_imageFile3", model.old_imageFile3);
            requestWriter.Write("\r\n--" + boundary + "\r\n");
            requestWriter.Write(requestFormat, "old_imageFile4", model.old_imageFile4);
            requestWriter.Write("\r\n--" + boundary + "\r\n");
            requestWriter.Write(requestFormat, "old_imageFile5", model.old_imageFile5);
            requestWriter.Write("\r\n--" + boundary + "\r\n");
            requestWriter.Write(requestFormat, "old_imageFile6", model.old_imageFile6);
            requestWriter.Write("\r\n--" + boundary + "\r\n");
            requestWriter.Write(requestFormat, "old_imageFile7", model.old_imageFile7);
            requestWriter.Write("\r\n--" + boundary + "\r\n");
            requestWriter.Write(requestFormat, "old_imageFile8", model.old_imageFile8);
            requestWriter.Write("\r\n--" + boundary + "\r\n");
            requestWriter.Write(requestFormat, "old_imageFile9", model.old_imageFile9);
            requestWriter.Write("\r\n--" + boundary + "\r\n");
            requestWriter.Write(requestFormat, "old_imageFile10", model.old_imageFile10);
            requestWriter.Write("\r\n--" + boundary + "\r\n");
            requestWriter.Write(requestFormat, "bdIpAddress", model.bdIpAddress);
            requestWriter.Write("\r\n--" + boundary + "\r\n");
            requestWriter.Write(requestFormat, "bdTitle", model.bdTitle);
            requestWriter.Write("\r\n--" + boundary + "\r\n");
            requestWriter.Write(requestFormat, "bdName", model.bdName);
            requestWriter.Write("\r\n--" + boundary + "\r\n");
            requestWriter.Write(requestFormat, "bdPassword", model.bdPassword);
            requestWriter.Write("\r\n--" + boundary + "\r\n");
            requestWriter.Write(requestFormat, "bdEmail", model.bdEmail);
            requestWriter.Write("\r\n--" + boundary + "\r\n");
            requestWriter.Write(requestFormat, "bdPhone", model.bdPhone);
            requestWriter.Write("\r\n--" + boundary + "\r\n");
            requestWriter.Write(requestFormat, "bdType", model.bdType);
            requestWriter.Write("\r\n--" + boundary + "\r\n");
            requestWriter.Write(requestFormat, "bdLocation", model.bdLocation);
            requestWriter.Write("\r\n--" + boundary + "\r\n");
            requestWriter.Write(requestFormat, "bdPrice", model.bdPrice);
            requestWriter.Write("\r\n--" + boundary + "\r\n");
            requestWriter.Write(requestFormat, "bdTag", model.bdTag);
            requestWriter.Write("\r\n--" + boundary + "\r\n");
            requestWriter.Write(requestFormat, "chk_image", model.chk_image);
            requestWriter.Write("\r\n--" + boundary + "\r\n");
            requestWriter.Write(requestFormatWithFile, "photo1", model.photo1, "");
            requestWriter.Write("\r\n--" + boundary + "\r\n");
            requestWriter.Write(requestFormatWithFile, "photo2", model.photo2, "");
            requestWriter.Write("\r\n--" + boundary + "\r\n");
            requestWriter.Write(requestFormatWithFile, "photo3", model.photo3, "");
            requestWriter.Write("\r\n--" + boundary + "\r\n");
            requestWriter.Write(requestFormatWithFile, "photo4", model.photo4, "");
            requestWriter.Write("\r\n--" + boundary + "\r\n");
            requestWriter.Write(requestFormatWithFile, "photo5", model.photo5, "");
            requestWriter.Write("\r\n--" + boundary + "\r\n");
            requestWriter.Write(requestFormatWithFile, "photo6", model.photo6, "");
            requestWriter.Write("\r\n--" + boundary + "\r\n");
            requestWriter.Write(requestFormatWithFile, "photo7", model.photo7, "");
            requestWriter.Write("\r\n--" + boundary + "\r\n");
            requestWriter.Write(requestFormatWithFile, "photo8", model.photo8, "");
            requestWriter.Write("\r\n--" + boundary + "\r\n");
            requestWriter.Write(requestFormatWithFile, "photo9", model.photo9, "");
            requestWriter.Write("\r\n--" + boundary + "\r\n");
            requestWriter.Write(requestFormatWithFile, "photo10", model.photo10, "");
            requestWriter.Write("\r\n--" + boundary + "\r\n");
            requestWriter.Write(requestFormat, "bdDescription", model.bdDescription);
            requestWriter.Write("\r\n--" + boundary + "\r\n");
         }
         catch (Exception e){
            TraceLog.Instance.WriteLine(string.Format("Error: {0}", e.Message));
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

      private void CheckDataFile(string fileName)
      {
         if (!System.IO.File.Exists(fileName))
         {
            Console.WriteLine("There is no " + fileName + " in the directory.");
            return;
         }
      }

      private List<BoardIDModel> ReadDataFile(string fileName)
      {
         Excel.Application xlApp;
         Excel.Workbook xlWorkBook;
         Excel.Worksheet xlWorkSheet;
         Excel.Range range;

         List<BoardIDModel> reqList = new List<BoardIDModel>();

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
               BoardIDModel model = new BoardIDModel();
               for (int i = 1; i <= rowNo; i++)
               {
                  if (array[i, j] != null)
                  {
                     if (array[i, 1].ToString() == "boardId")
                        model.boardId = array[i, j].ToString();
                     if (array[i, 1].ToString() == "bdId")
                        model.bdId = array[i, j].ToString();
                     if (array[i, 1].ToString() == "bdPassword")
                        model.bdPassword = array[i, j].ToString();
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
         catch (Exception e){
            TraceLog.Instance.WriteLine(string.Format("Error: {0}", e.Message));
         }

         return reqList;
      }

      private void releaseObject(object obj)
      {
         try
         {
            System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
            obj = null;
         }
         catch (Exception e){
            obj = null;
            TraceLog.Instance.WriteLine(string.Format("Error: {0}", e.Message));
         }
         finally
         {
            GC.Collect();
         }
      }
   }
}