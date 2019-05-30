using Microsoft.Dynamics365.UIAutomation.Api.UCI;
using Microsoft.Dynamics365.UIAutomation.Browser;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Security;

namespace Microsoft.Dynamics365.UIAutomation.Sample.UCI.Create
{
    [TestClass]
    public class CreateForecast
    {
        private readonly SecureString _username = System.Configuration.ConfigurationManager.AppSettings["OnlineUsername"].ToSecureString();
        private readonly SecureString _password = System.Configuration.ConfigurationManager.AppSettings["OnlinePassword"].ToSecureString();
        private readonly Uri _xrmUri = new Uri(System.Configuration.ConfigurationManager.AppSettings["OnlineCrmUrl"].ToString());

        [TestMethod]
        public void UCITestCreateForecast()
        {
            WebClient client = new WebClient(TestSettings.Options);
            #region xrmAPP
            using (XrmApp xrmApp = new XrmApp(client))
            {
                xrmApp.OnlineLogin.Login(_xrmUri, _username, _password);
                xrmApp.ThinkTime(25000);

                xrmApp.Navigation.OpenApp(UCIAppName.Sourcing);

                xrmApp.Navigation.OpenSubArea("Sourcing", "Forecasts");
                //xrmApp.ThinkTime(1000);

                xrmApp.CommandBar.ClickCommand("New");

                xrmApp.ThinkTime(5000);

                //Set all the required props
                xrmApp.Entity.SetValue("src_name", "BVT" + TestSettings.GetRandomString(5, 15));
                //Test code to search and set look up 
                // LookupItem parentAccountId = new LookupItem { Name = "src_businessowner" };
                //xrmApp.Lookup.Search(parentAccountId, "sukar@microsoft.com");

                //xrmApp.ThinkTime(1000);
                //xrmApp.Lookup.OpenRecord(0);

                //xrmApp.Entity.SelectLookup(new LookupItem { Name = "src_businessowner" });

                //xrmApp.Entity.SetValue("src_businessowner", "Karlton Zuna");
                //LookupItem item = new LookupItem { Name = "src_businessowner", Index = 0 };
                //xrmApp.Entity.SetValue(item);

                //LookupItem parentAccountId = new LookupItem { Name = "src_businessowner", Value = "Khean Hooi Tanemura" };
                // xrmApp.Entity.SelectLookup(parentAccountId);
                xrmApp.Entity.SetValue("src_businessowner", "sukar@microsoft.com");
                //xrmApp.Lookup.Search(parentAccountId, "sukar@microsoft.com");

                xrmApp.Entity.SetValue("src_needbydate", "4/15/2019");
                //xrmApp.Entity.SetValue("src_needbydate", DateTime.Today, "MM/dd/yyyy");
                xrmApp.Entity.SetValue("src_description", "BVT");

                //xrmApp.ThinkTime(1000);
                xrmApp.Entity.SetValue(new LookupItem { Name = "src_categorygroup", Value = "Events" });
                xrmApp.Entity.SetValue(new LookupItem { Name = "src_categoryname", Value = "Events Services" });
                //xrmApp.Entity.SetValue(new LookupItem { Name = "src_categorygroup", Index = 0 });
                //xrmApp.Entity.SetValue("src_categoryname", "Event Services");
                xrmApp.Entity.SetValue("src_spend", "45678");
                //xrmApp.ThinkTime(1000);
                //xrmApp.Entity.SetValue("src_sourcingteaminvolvement", "Don't know");
                //xrmApp.Entity.SetValue("src_likelihood", "0-25%");
                //xrmApp.ThinkTime(1000);
                //xrmApp.Entity.SetValue("src_forecasttype", "Contract Negotiation");
                //xrmApp.Entity.SetValue("src_projectstartdate", "4/15/2019");
                //// xrmApp.Entity.SetValue("Description", "BVT");

                xrmApp.ThinkTime(50000);


                xrmApp.Entity.Save();

            }
            #endregion




        }

    }
}
