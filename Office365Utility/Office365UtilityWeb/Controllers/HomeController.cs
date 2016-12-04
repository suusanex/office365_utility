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

namespace Office365UtilityWeb.Controllers
{
    public class HomeController : Controller
    {
        public ActionResult Index()
        {

            ViewBag.Url_GetAllGroupsAndUsersList = string.Join("&",
                "https://login.microsoftonline.com/common/oauth2/authorize?response_type=code",
                $"client_id={ClientId}",
                "resource=https%3a%2f%2fgraph.microsoft.com%2f",
                $"redirect_uri={HttpUtility.UrlEncode($"http://{new Uri(Request.Url.AbsoluteUri).Authority}/Home/AuthResult")}"
                ); 

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


        string ClientId { get { return ConfigurationManager.AppSettings["ClientId"]; } }
        string ClientSecret { get { return ConfigurationManager.AppSettings["ClientSecret"]; } }
        string TenantName { get { return ConfigurationManager.AppSettings["TenantName"]; } }

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

                var O365GetTokenPOSTURL = "https://login.microsoftonline.com/common/oauth2/token";
                var O365GetTokenPOSTData = string.Join("&",
                    "grant_type=authorization_code",
                    $"code={code}",
                    $"client_id={ClientId}",
                    $"client_secret={HttpUtility.UrlEncode(ClientSecret)}",
                    $"redirect_uri={HttpUtility.UrlEncode(RedirectUrl)}",
                    "resource=https%3a%2f%2fgraph.microsoft.com%2f"
                    );

                TraceDebug(O365GetTokenPOSTData);

                Models.Office365TokenResponseJson TokenData;
                TokenData = GetTokenData(O365GetTokenPOSTURL, O365GetTokenPOSTData);

                var CSVBuf = new StringBuilder();
                CSVBuf.AppendLine(string.Join(",", "Group Mail", "Group Type", "User Mail"));

                Models.MicrosotfGraph_GroupsJson Groups;
                Groups = GetAllGroups(TokenData.access_token);

                foreach (var group in Groups.value)
                {
                    //メアドのないAdminAgentsというグループがあり、ML取得には無関係なのでスキップ
                    if (string.IsNullOrEmpty(group.mail))
                    {
                        continue;
                    }

                    Models.MicrosoftGraph_UsersJson Users;

                    Users = GetUsersFromGroup(TokenData.access_token, group.id);


                    foreach (var user in Users.value)
                    {
                        CSVBuf.AppendLine(string.Join(",", group.mail, string.Join("/", group.groupTypes), user.mail));
                    }
                }

                ViewBag.GroupAndUsersCSV = CSVBuf.ToString();
                return View("Index");
            }
            catch (WebException ex)
            {
                TraceDebug(ex.ToString());

                try
                {
                    string resData;
                    using (var res = new System.IO.StreamReader(ex.Response.GetResponseStream()))
                    {
                        resData = res.ReadToEnd();
                    }
                    TraceDebug(resData);
                }
                catch (Exception ex2)
                {
                    TraceDebug(ex2.ToString());

                }

                throw;
            }
            catch (Exception ex)
            {
                TraceDebug(ex.ToString());
                throw;
            }
        }

        private Models.MicrosotfGraph_GroupsJson GetAllGroups(string access_token)
        {
            Models.MicrosotfGraph_GroupsJson Groups;
            using (var wc = new WebClient())
            {
                wc.Encoding = Encoding.UTF8;
                wc.Headers.Add("Accept", "application/json");
                wc.Headers.Add("Authorization", $"Bearer {access_token}");

                var getURL = $"https://graph.microsoft.com/v1.0/{TenantName}/groups";

                using (var readstream = wc.OpenRead(getURL))
                {
                    using (var strReadStream = new StreamReader(readstream, Encoding.UTF8))
                    {
                        var GetStr = strReadStream.ReadToEnd();
                        TraceDebug(GetStr);
                        Groups = JsonConvert.DeserializeObject<Models.MicrosotfGraph_GroupsJson>(GetStr);
                    }
                }

            }

            return Groups;
        }

        private Models.MicrosoftGraph_UsersJson GetUsersFromGroup(string access_token, string groupId)
        {
            Models.MicrosoftGraph_UsersJson Users;
            using (var wc = new WebClient())
            {
                wc.Encoding = Encoding.UTF8;
                wc.Headers.Add("Accept", "application/json");
                wc.Headers.Add("Authorization", $"Bearer {access_token}");


                var getURL = $"https://graph.microsoft.com/v1.0/{TenantName}/groups/{groupId}/members";

                using (var readstream = wc.OpenRead(getURL))
                {
                    using (var strReadStream = new StreamReader(readstream, Encoding.UTF8))
                    {
                        var GetStr = strReadStream.ReadToEnd();
                        TraceDebug(GetStr);
                        Users = JsonConvert.DeserializeObject<Models.MicrosoftGraph_UsersJson>(GetStr);
                    }
                }

            }

            return Users;
        }

        private Models.Office365TokenResponseJson GetTokenData(string O365GetTokenPOSTURL, string O365GetTokenPOSTData)
        {
            Models.Office365TokenResponseJson TokenData;
            try
            {
                using (var WebCli = new WebClient())
                {
                    WebCli.Encoding = Encoding.UTF8;
                    WebCli.Headers.Add("Content-Type", "application/x-www-form-urlencoded");

                    string resData = WebCli.UploadString(O365GetTokenPOSTURL, O365GetTokenPOSTData);

                    TraceDebug(resData.ToString());

                    TokenData = JsonConvert.DeserializeObject<Models.Office365TokenResponseJson>(resData);

                }
            }
            catch (WebException ex)
            {
                TraceDebug(ex.ToString());

                try
                {
                    string resData;
                    using (var res = new System.IO.StreamReader(ex.Response.GetResponseStream()))
                    {
                        resData = res.ReadToEnd();
                    }
                    var getResultObj = JsonConvert.DeserializeObject<Models.Office365TokenErrorJson>(resData);

                    TraceDebug(getResultObj.ToString());
                }
                catch (Exception ex2)
                {
                    TraceDebug(ex2.ToString());

                }

                throw;
            }

            return TokenData;
        }
    }
}