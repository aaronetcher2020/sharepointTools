using System;
using System.IO;
using System.Threading.Tasks;
using Azure.Identity;
using Microsoft.Graph;
using Microsoft.Graph.Models;
using Microsoft.Graph.Models.Security;
using Microsoft.Identity.Client;
using Newtonsoft.Json;

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
            foreach (var item in items.Value)
            {
                if (item.CreatedDateTime > filterDate)
                {
                    string path = currentFilePath + "/" + item.WebUrl.Replace(siteUrl + "Shared%20Documents/", "").ToString();
                    if (item.Folder != null)
                    {

                        if (System.IO.Directory.Exists(path))
                        { }
                        else
                        {
                            System.IO.Directory.CreateDirectory(path);
                            Console.WriteLine("Created the following Folder Path:" + path);
                        }
                        GetFolderData(driveId, item.Id);
                    }
                    else
                    {


                        if (System.IO.Directory.Exists(path))
                        { }
                        else
                        {
                            System.IO.Directory.CreateDirectory(path);
                        }


                        SaveFileStream(path + "/" + item.Name, graphClient.Drives[driveId].Items[item.Id].Content.GetAsync().Result);
                        Console.WriteLine("Saved Files to the following Folder: " + path + " with filename " + item.Name);

                    }
                }
                else
                {
                    Console.WriteLine("Item before start date");
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

}