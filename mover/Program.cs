using System;
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

        static async Task Main(string[] args)
        {
            //Sharepoint URL form run input
            //https://landerholmlaw.sharepoint.com/sites/
            siteUrl = args[0].ToString();
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
                specificFolder = args[4];
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

            //	Id	"landerholmlaw.sharepoint.com,7df7751a-9d60-4711-86e5-674d1f236baf,642dd668-9da8-4b9c-b3eb-e5bdbcc0f534"	
            var site = await graphClient.Sites[id].GetAsync();
            var drives = await graphClient.Sites[site.Id].Drives.GetAsync(); ;
            var d2 = await graphClient.Drives[drives.Value[0].Id].Root.GetAsync();
           // var items = await graphClient.Drives[drives.Value[0].Id.ToString()].Items[d2.Id].Children.GetAsync();
            GetFolderData(drives.Value[0].Id.ToString(), d2.Id);




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
        
        public static void GetFolderData(string driveId, string itemsId)
        {
            var items = graphClient.Drives[driveId].Items[itemsId].Children.GetAsync().Result;
           // var targetedPath = 
            foreach (var item in items.Value)
            {
                if (item.CreatedDateTime > filterDate)
                {
                    Console.WriteLine(siteUrl);
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



                                Console.WriteLine("Created the following Folder Path:" + path);
                            }
                        }
                        
                        GetFolderData(driveId, item.Id);
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
                            SaveFileStream(path, graphClient.Drives[driveId].Items[item.Id].Content.GetAsync().Result);
                            TimeZoneInfo timezone = TimeZoneInfo.FindSystemTimeZoneById("Pacific Standard Time");
                            System.IO.File.SetCreationTime(path, TimeZoneInfo.ConvertTime(item.CreatedDateTime.Value.DateTime, timezone).AddHours(-7));
                            System.IO.File.SetLastWriteTime(path, TimeZoneInfo.ConvertTime(item.LastModifiedDateTime.Value.DateTime, timezone).AddHours(-7));
                            Console.WriteLine("Saved Files to the following Folder: " + path + " with filename " + item.Name);
                        }

                    }
                }
                else
                {
                    Console.WriteLine("Item before start date");
                }
            }
            updateTimeStamps(driveId, itemsId);
            

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
                    }
                    catch (Exception e) { }
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