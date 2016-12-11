using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Web;
using System.Diagnostics;
using System.Net;
using System.IO;

using Newtonsoft.Json;

namespace Office365UtilityLib
{
    public class MicrosoftGraph_v1
    {
        public MicrosoftGraph_v1(string ClientId, string ClientSecret, string TenantName)
        {
            m_ClientId = ClientId;
            m_ClientSecret = ClientSecret;
            m_TenantName = TenantName;
        }

        public string GenerateAuthStartUrl(string RedirectUrl)
        {
            return string.Join("&",
                "https://login.microsoftonline.com/common/oauth2/authorize?response_type=code",
                $"client_id={m_ClientId}",
                "resource=https%3a%2f%2fgraph.microsoft.com%2f",
                $"redirect_uri={HttpUtility.UrlEncode(RedirectUrl)}"
                );

        }


        public void UpdateTokenData(string AuthCode, string RedirectUrl)
        {
            var O365GetTokenPOSTURL = "https://login.microsoftonline.com/common/oauth2/token";
            var O365GetTokenPOSTData = string.Join("&",
                "grant_type=authorization_code",
                $"code={AuthCode}",
                $"client_id={m_ClientId}",
                $"client_secret={HttpUtility.UrlEncode(m_ClientSecret)}",
                $"redirect_uri={HttpUtility.UrlEncode(RedirectUrl)}",
                "resource=https%3a%2f%2fgraph.microsoft.com%2f"
                );

            Office365API.TokenResponse TokenData;
            try
            {
                using (var WebCli = new WebClient())
                {
                    WebCli.Encoding = Encoding.UTF8;
                    WebCli.Headers.Add("Content-Type", "application/x-www-form-urlencoded");

                    string resData = WebCli.UploadString(O365GetTokenPOSTURL, O365GetTokenPOSTData);

                    TraceDebug(resData.ToString());

                    TokenData = JsonConvert.DeserializeObject<Office365API.TokenResponse>(resData);

                }
            }
            catch (WebException ex)
            {
                TraceWebException(ex);

                throw;
            }

            m_TokenData = TokenData;
        }


        public MicrosoftGraph.Groups GetAllGroups()
        {
            string access_token = m_TokenData.access_token;

            MicrosoftGraph.Groups Groups;
            try
            {

                using (var wc = new WebClient())
                {
                    wc.Encoding = Encoding.UTF8;
                    wc.Headers.Add("Accept", "application/json");
                    wc.Headers.Add("Authorization", $"Bearer {access_token}");

                    var getURL = $"https://graph.microsoft.com/v1.0/{m_TenantName}/groups";

                    using (var readstream = wc.OpenRead(getURL))
                    {
                        using (var strReadStream = new StreamReader(readstream, Encoding.UTF8))
                        {
                            var GetStr = strReadStream.ReadToEnd();
                            TraceDebug(GetStr);
                            Groups = JsonConvert.DeserializeObject<MicrosoftGraph.Groups>(GetStr);
                        }
                    }

                }
            }
            catch (WebException ex)
            {
                TraceWebException(ex);

                throw;
            }

            return Groups;
        }

        public MicrosoftGraph.Users GetUsersFromGroup(string groupId)
        {
            string access_token = m_TokenData.access_token;

            MicrosoftGraph.Users Users;

            try
            {

                using (var wc = new WebClient())
                {
                    wc.Encoding = Encoding.UTF8;
                    wc.Headers.Add("Accept", "application/json");
                    wc.Headers.Add("Authorization", $"Bearer {access_token}");


                    var getURL = $"https://graph.microsoft.com/v1.0/{m_TenantName}/groups/{groupId}/members";

                    using (var readstream = wc.OpenRead(getURL))
                    {
                        using (var strReadStream = new StreamReader(readstream, Encoding.UTF8))
                        {
                            var GetStr = strReadStream.ReadToEnd();
                            TraceDebug(GetStr);
                            Users = JsonConvert.DeserializeObject<MicrosoftGraph.Users>(GetStr);
                        }
                    }

                }
            }
            catch (WebException ex)
            {
                TraceWebException(ex);

                throw;
            }

            return Users;
        }

        private void TraceWebException(WebException ex)
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
        }

        [Conditional("DEBUG")]
        void TraceDebug(string msg)
        {
            Debug.WriteLine(msg + Environment.NewLine);

        }

        internal string m_ClientId;
        internal string m_ClientSecret;
        internal string m_TenantName;
        Office365API.TokenResponse m_TokenData;
    }
}
