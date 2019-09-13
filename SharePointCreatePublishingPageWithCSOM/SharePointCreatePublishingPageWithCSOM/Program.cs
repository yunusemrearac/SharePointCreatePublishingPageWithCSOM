using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.Publishing;
using System;
using System.Linq;
using System.Net;

namespace SharePointCreatePublishingPageWithCSOM
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Please Enter Site Address: ");
            string siteUrl = Console.ReadLine();
            bool siteOnline = false;
            string siteUser = "";
            string sitePassword = "";
            string siteDomain = "";

            if (siteUrl.Contains(".sharepoint.com"))
            {
                Console.WriteLine("This site is SharePoint Online Right ? (Y/N): ");
                string OnlineAnswer = Console.ReadLine();

                if (OnlineAnswer.ToUpper() == "Y")
                {
                    siteOnline = true;
                }
            }

            if (siteOnline)
            {
                Console.WriteLine("Please Enter Site Login User Email Address: ");
                siteUser = Console.ReadLine();

                Console.WriteLine("Please Enter Site Login User Password: ");
                sitePassword = Console.ReadLine();
            }
            else
            {
                Console.WriteLine("Please Enter Site Login User Domain Name: ");
                siteDomain = Console.ReadLine();

                Console.WriteLine("Please Enter Site Login User Name: ");
                siteUser = Console.ReadLine();

                Console.WriteLine("Please Enter Site Login User Password: ");
                sitePassword = Console.ReadLine();
            }

            Console.WriteLine("Please Enter Master Page Gallery Name: ");
            string siteMasterPageGallery = Console.ReadLine();

            Console.WriteLine("Please Enter Create Publishing Page Content Type Display Name: ");
            string pageContentType = Console.ReadLine();

            Console.WriteLine("Please Enter Create Publishing Page Name: ");
            string pageName = Console.ReadLine();

            Console.WriteLine("Do You Have Add Page Field Value (Y/N): ");
            string addPageValuAnswer = Console.ReadLine();
            bool addPageValue = false;
            string addPageValueString = "";

            if (addPageValuAnswer.ToUpper() == "Y")
            {
                addPageValue = true;
                Console.WriteLine("Please Enter Page Field Data ( ':' Separator Using For Field Name And Value, ';' Separator Using For Each Field Data ): ");
                Console.WriteLine("For Example ---> Title:TestPage;Comments:TestPage Comments ");
                Console.WriteLine("");
                addPageValueString = Console.ReadLine();
            }

            Console.WriteLine("Please Wait Start Create Process....");

            ClientContext clientContext = null;
            PublishingPage publishingPage = null;
            Web web = null;

            try
            {
                clientContext = new ClientContext(siteUrl);

                if (siteOnline)
                {
                    clientContext.AuthenticationMode = ClientAuthenticationMode.Default;

                    string password = sitePassword;
                    System.Security.SecureString passwordChar = new System.Security.SecureString();
                    foreach (char ch in password)
                        passwordChar.AppendChar(ch);

                    clientContext.Credentials = new SharePointOnlineCredentials(siteUser, passwordChar);
                }
                else
                {
                    clientContext.Credentials = new NetworkCredential(siteUser, sitePassword, siteDomain);
                }

                web = clientContext.Web;
                clientContext.Load(web);
                clientContext.ExecuteQuery();
            }
            catch (Exception ex)
            {
                Console.WriteLine("This Site Connect Error Please Check Your Information And Try Again");
                Console.WriteLine("Error Message: " + ex.Message);
                Console.ReadLine();
                Environment.Exit(0);
            }

            try
            {
                Console.WriteLine("This Site Connection Success....");
                Console.WriteLine("Please Wait Create Publishing Page.....");

                List publishingLayouts = clientContext.Site.RootWeb.Lists.GetByTitle(siteMasterPageGallery);
                ListItemCollection allItems = publishingLayouts.GetItems(CamlQuery.CreateAllItemsQuery());
                clientContext.Load(allItems, items => items.Include(item => item.DisplayName).Where(obj => obj.DisplayName == pageContentType));
                clientContext.ExecuteQuery();
                ListItem layout = allItems.Where(x => x.DisplayName == pageContentType).FirstOrDefault();
                clientContext.Load(layout);

                PublishingWeb publishingWeb = PublishingWeb.GetPublishingWeb(clientContext, web);
                clientContext.Load(publishingWeb);

                PublishingPageInformation publishingPageInfo = new PublishingPageInformation();
                publishingPageInfo.Name = pageName.Contains(".aspx") ? pageName : pageName + ".aspx";
                publishingPageInfo.PageLayoutListItem = layout;

                publishingPage = publishingWeb.AddPublishingPage(publishingPageInfo);

                clientContext.Load(publishingPage);
                clientContext.Load(publishingPage.ListItem.File);
                clientContext.ExecuteQuery();
            }
            catch (Exception ex)
            {
                Console.WriteLine("This During Publishing Page Error Please Check Your Information And Try Again");
                Console.WriteLine("Error Message: " + ex.Message);
                Console.ReadLine();
                Environment.Exit(0);
            }

            Console.WriteLine("this Publishing Page Create Success....");

            if (addPageValue)
            {
                try
                {
                    Console.WriteLine("Please Wait Add Field Data Publishing Page....");

                    ListItem listItem = publishingPage.ListItem;

                    string[] dataArray = addPageValueString.Split(';');

                    foreach (string data in dataArray)
                    {
                        listItem[data.Split(':')[0]] = data.Split(':')[1];
                    }

                    listItem.Update();

                    publishingPage.ListItem.File.CheckIn(string.Empty, CheckinType.MajorCheckIn);
                    publishingPage.ListItem.File.Publish(string.Empty);

                    clientContext.Load(publishingPage);

                    clientContext.ExecuteQuery();

                    Console.WriteLine("Tihs Publishing Page Add Field Data Success....");
                }
                catch (Exception ex)
                {
                    Console.WriteLine("This During Publishing Page Add Field Data Error Please Check Your Information And Try Again");
                    Console.WriteLine("Error Message: " + ex.Message);
                    Console.ReadLine();
                    Environment.Exit(0);
                }
            }

            Console.WriteLine("All Process Complete Success...");
            Console.ReadLine();
            Environment.Exit(0);
        }
    }
}
