using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using System.Text;
using System.Net;
using System.IO;
using System.Diagnostics;
using System.Configuration;

using Newtonsoft.Json;

using Office365UtilityLib;

namespace Office365UtilityWeb.Controllers
{
    internal static class Office365API
    {
        static Office365API()
        {
            Graph = new MicrosoftGraph_v1(
                ConfigurationManager.AppSettings["ClientId"],
                ConfigurationManager.AppSettings["ClientSecret"],
                ConfigurationManager.AppSettings["TenantName"]
                );
        }

        internal static MicrosoftGraph_v1 Graph;
    }

    public class HomeController : Controller
    {
        

        public ActionResult Index()
        {
            var RedirectUrl = $"http://{new Uri(Request.Url.AbsoluteUri).Authority}/Home/AuthResult";
            ViewBag.Url_GetAllGroupsAndUsersList =
                Office365API.Graph.GenerateAuthStartUrl(RedirectUrl); 
            
            return View();
        }

        public ActionResult About()
        {
            ViewBag.Message = "Your application description page.";

            return View();
        }

        public ActionResult Contact()
        {
            ViewBag.Message = "Your contact page.";

            return View();
        }



        [System.Diagnostics.Conditional("DEBUG")]
        void TraceDebug(string msg)
        {
            System.Diagnostics.Debug.WriteLine(msg + Environment.NewLine);
            
        }

        [HttpGet]
        public ActionResult AuthResult(string code, string session_state)
        {
            try
            {
                TraceDebug(code);
                TraceDebug(session_state);


                var RedirectUrl = new Uri(new Uri(Request.Url.AbsoluteUri), "AuthResult").AbsoluteUri;

                Office365API.Graph.UpdateTokenData(code, RedirectUrl);

                var CSVBuf = new StringBuilder();
                CSVBuf.AppendLine(string.Join(",", "Group Mail", "Group Type", "User Mail"));

                var Groups = Office365API.Graph.GetAllGroups();

                foreach (var group in Groups.value)
                {
                    //メアドのないAdminAgentsというグループがあり、ML取得には無関係なのでスキップ
                    if (string.IsNullOrEmpty(group.mail))
                    {
                        continue;
                    }


                    var Users = Office365API.Graph.GetUsersFromGroup(group.id);


                    foreach (var user in Users.value)
                    {
                        CSVBuf.AppendLine(string.Join(",", group.mail, string.Join("/", group.groupTypes), user.mail));
                    }
                }

                ViewBag.GroupAndUsersCSV = CSVBuf.ToString();
                return View("Index");
            }
            catch (Exception ex)
            {
                TraceDebug(ex.ToString());
                throw;
            }
        }


    }
}