using System;
using System.IO;
using Azure.Identity;
using Microsoft.Graph;
using Microsoft.Graph.Models;
using PnP.Core.Model.SharePoint;


namespace GraphODataDemo
{
    class Program
    {
        public static GraphServiceClient graphClient { get; set; }
        public static string siteUrl { get; set; }
        public static string siteName { get; set; }
        public static string targetFolder { get; set; }
        public static DateTime filterDate { get; set; }
        public static string currentFilePath { get; set; }
        public static string specificFolder { get; set; }
        public static int folderCount { get; set; }
        public static int fileCount { get; set; }

        static async Task Main(string[] args)
        {
            //Sharepoint URL form run input
            //https://landerholmlaw.sharepoint.com/sites/
            siteUrl = args[0].ToString();
            folderCount = 0;
            fileCount = 0;
            //DevA/
            siteName = args[1].ToString();
            targetFolder = args[2].ToString();
            if (args[3] != null)
            {
                filterDate = Convert.ToDateTime(args[3]);
            }
            else
            {
                filterDate = DateTime.MinValue;
            }
            if(args.Length> 4)
            {
                specificFolder = System.Web.HttpUtility.UrlDecode(args[4]);
                //specificFolder = args[4];
            }
            else
            {
                specificFolder = string.Empty;
            }


            graphClient = SetupClient();
            currentFilePath =targetFolder;
            currentFilePath = currentFilePath + "/Shared Documents";
            if (System.IO.Directory.Exists(currentFilePath))
            { }
            else
            {
                    System.IO.Directory.CreateDirectory(currentFilePath);
          
            }
            siteUrl = siteUrl + siteName;

            var filteredSite = await graphClient.Sites.GetAllSites.GetAsync(requestConfig =>
            {
                requestConfig.QueryParameters.Filter =
                    "webUrl eq '"+siteUrl+"'";
            });

            var id = filteredSite.Value.First().Id.Split(',')[1].ToString();
            var site = await graphClient.Sites[id].GetAsync();
            var drives = await graphClient.Sites[site.Id].Drives.GetAsync(); ;
            var d2 = await graphClient.Drives[drives.Value[0].Id].Root.GetAsync();
            
          var i = await GetFolderData(drives.Value[0].Id.ToString(), d2.Id);
            updateTimeStamps(drives.Value[0].Id, d2.Id);



            Console.WriteLine("FINISHED PROCESS");

            
        }
        private static GraphServiceClient SetupClient()
        {
            string[] scopes = { "https://graph.microsoft.com/.default" };
            string clientId = "2e9732e9-086b-4074-8344-740d521f0b23";
            string secret = "Yh~8Q~hy4yKoiaZFx1gpKLhYW.FwmTvz48~SscAx";
            string tenant = "1c55f0ab-4e28-451b-bd29-785d97b143ab";
            ClientSecretCredential clientSecretCredential = new ClientSecretCredential(tenant, clientId, secret);

            return new GraphServiceClient(clientSecretCredential, scopes);
        }

        public static async Task<int> GetFolderData(string driveId, string itemsId )
        {
            var items = graphClient.Drives[driveId].Items[itemsId].Children.GetAsync().Result;
           // var targetedPath = 
            foreach (var item in items.Value)
            {
                if (item.CreatedDateTime > filterDate)
                {
                   // Console.WriteLine(siteUrl);
                    string pt = System.Web.HttpUtility.UrlDecode(item.WebUrl);

                    string ck = item.WebUrl;
                    string ck2 = ck.Replace(siteUrl, "");
                    string replaced = ck2.Replace("/Shared%20Documents", "");
                    string path = currentFilePath + replaced;
                    path = System.Web.HttpUtility.UrlDecode(path);
                    if (item.Folder != null)
                    {
                        if (specificFolder.Equals(string.Empty)|| item.WebUrl.Contains(specificFolder))
                        {

                            if (System.IO.Directory.Exists(path))
                            { }
                            else
                            {
                                System.IO.Directory.CreateDirectory(path);


                                //folderCount++;
                                Console.WriteLine("Created the following Folder " + folderCount.ToString() + " Path:" + path + " folder new" + item.Name); 
                            }
                        }
                        Console.WriteLine("Starting a new sub folder at path:" + path);
                        var i2 = GetFolderData(driveId, item.Id);
                    }
                    else
                    {
                        if (item.WebUrl.Contains(specificFolder) || specificFolder.Equals(string.Empty))
                        {
                            if (System.IO.File.Exists(path))
                            {
                                System.IO.File.Delete(path);
                            }
                            else
                            {

                            }
                            fileCount++;
                            SaveFileStream(path, graphClient.Drives[driveId].Items[item.Id].Content.GetAsync().Result);
                            TimeZoneInfo timezone = TimeZoneInfo.FindSystemTimeZoneById("Pacific Standard Time");
                            System.IO.File.SetCreationTime(path, TimeZoneInfo.ConvertTime(item.CreatedDateTime.Value.DateTime, timezone).AddHours(-7));
                            System.IO.File.SetLastWriteTime(path, TimeZoneInfo.ConvertTime(item.LastModifiedDateTime.Value.DateTime, timezone).AddHours(-7));
                            Console.WriteLine("Saved Files "+fileCount.ToString()+" to the following Folder: " + path + " with filename " + item.Name);
                        }

                    }
                }
                else
                {
                    Console.WriteLine("Item before start date");
                }
            }
            return 0;
            
            

        }
        public static void updateTimeStamps (string driveId, string itemsId)
        {
            var items = graphClient.Drives[driveId].Items[itemsId].Children.GetAsync().Result;
            foreach (var item in items.Value)
            {
                if (item.Folder != null)
                {
                    //
                    string ck = item.WebUrl;
                    string ck2 = ck.Replace(siteUrl, "");
                    string replaced = ck2.Replace("/Shared%20Documents", "");
                    string path = currentFilePath + replaced;
                    path = System.Web.HttpUtility.UrlDecode(path);
                    updateTimeStamps(driveId, item.Id);
                    TimeZoneInfo timezone = TimeZoneInfo.FindSystemTimeZoneById("Pacific Standard Time");
                    try
                    {

                        System.IO.Directory.SetCreationTime(path, TimeZoneInfo.ConvertTime(item.CreatedDateTime.Value.DateTime, timezone).AddHours(-7));
                        System.IO.Directory.SetLastWriteTime(path, TimeZoneInfo.ConvertTime(item.LastModifiedDateTime.Value.DateTime, timezone).AddHours(-7));
                        Console.WriteLine("Updated Timestamp(s) on folder:" + item.Name);
                    }
                    catch (Exception e) {
                        var et = e;
                    }
                }
            }

        }
        public static void SaveFileStream(String path, Stream stream)
        {
            var fileStream = new FileStream(path, FileMode.Create, FileAccess.Write);
           
            stream.CopyTo(fileStream);
            fileStream.Dispose();
        }
    }
    public class item
    {
        public string path { get; set; }
        public DateTime createTime { get; set; }
        public DateTime modifiedTime { get; set; }
    }

}