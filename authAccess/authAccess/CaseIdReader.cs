using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Text.RegularExpressions;
using System.Net.Http;
using NPOI.HSSF.UserModel;
using static NPOI.XSSF.UserModel.Helpers.ColumnHelper;
using Nancy.Helpers;
using Newtonsoft.Json;

namespace authAccess
{
    class CaseIdReader
    {

        public  async Task<List<Content>> CaseNumberReaderAsync()
        {
           
            int j = 0;
            List<Content> writeCon = new List<Content>();
            //List<string> id = new List<string>();
            // List<DateTime> time = new List<DateTime>();
            var mails = OutlookEmails.ReadMailItems();

            foreach (var mail in mails)
            {
                String pattern1 = @"[1]\d{14}";
                String CaseReader = mail.EmailSubject;
                Match match1 = Regex.Match(CaseReader, pattern1, RegexOptions.IgnoreCase);
                if (match1.Success)
                {
                    Content data = new Content() { caseId = match1.Value, alias = mail.EmailTo };
                    data.time = mail.EmailDate;
                    data.iDate = mail.EmailDate.ToString("MM/dd");
                    data.iTime = mail.EmailDate.ToString("HH:mm");
                   // data.iDate = ;
                    String pattern2 = @"Task";
                    Match match2 = Regex.Match(CaseReader, pattern2, RegexOptions.IgnoreCase);
                    if (match2.Success)
                        data.isTask = "Collaboration Task";
                    else
                        data.isTask = "Case";
                    StringBuilder ser = new StringBuilder();
                    ser.Append(mail.EmailSubject);
                    ser.Append(mail.EmailBody);
                    string pattern3 = @"\s[A|B|C]\s";
                    Match match3 = Regex.Match(ser.ToString(), pattern3, RegexOptions.IgnoreCase);
                    if (match3.Success)
                        data.severity = match3.Value;
                    else
                        data.severity = "";

                    string json = await ReadApi(match1.Value);
                   // TokenInfoText.Text += $"Token: {sd.AccessToken}" + Environment.NewLine;
                   // TokenInfoText.Text += $"JSON: {ReadApi}" + Environment.NewLine;
                    dynamic dobj = JsonConvert.DeserializeObject<dynamic>(json);
                    data.Name = dobj["AgentId"].ToString();
                    data.Topic = dobj["Title"].ToString();
                    data.SupportCountry = dobj["SupportCountry"].ToString();
                    data.ServiceLevel = dobj["EntitlementInformation"]["ServiceLevel"].ToString();

                    writeCon.Add(data);
                    //id.Add(match1.Value);
                    //time.Add(mail.EmailDate);

                }
            }
            //Sort the numbers from the oldest to the lastest.
            // id.Sort();
            //remove the dulplicate numbers
            // List<string> id1=id.Distinct().ToList();
            // time.ToString();
            // Console.WriteLine(string.Join("\n", time)); 
            //  return id1;
            List<Content> nonDuplicateList = new List<Content>();
            foreach (Content mem in writeCon)
            {
                if (nonDuplicateList.Exists(x => x.caseId == mem.caseId) == false)
                {
                    nonDuplicateList.Add(mem);
                }
            }
            /* var sortedData =
               (from s in nonDuplicateList
                select new
                {
                    s.caseId,
                    s.time,
                    s.alias,
                    s.severity
                }).Distinct().OrderBy(x => x.caseId).ToList();
             foreach (var i in sortedData)
             {
                 Console.WriteLine("caseId:   " + i.caseId + "          " + "SentTime:   " + i.time + "          " + "Alias:   " + i.alias);

             }*/
            return nonDuplicateList;
        }
        public async Task<string> GetHttpContentWithTokenAsync(string url, string token)
        {
            var httpClient = new HttpClient();
            HttpResponseMessage response;
            try
            {
                var request = new HttpRequestMessage(HttpMethod.Get, url);
                //Add the token in Authorization header
                request.Headers.Authorization = new System.Net.Http.Headers.AuthenticationHeaderValue("Bearer", token);
                response = await httpClient.SendAsync(request);
                var content = await response.Content.ReadAsStringAsync();
                return content;
            }
            catch (Exception ex)
            {
                return ex.ToString();
            }

        }
        public async Task<string> ReadApi(string caseId)
        {
            SDAuthLib sd = new SDAuthLib();
            sd.GetSDToken();
            string url = "https://api.support.microsoft.com/v2/cases/";
            StringBuilder APIurl = new StringBuilder(url);
            APIurl.Append(String.Format("{0}", HttpUtility.HtmlEncode(caseId)));
            APIurl.Append(String.Format("{0}", HttpUtility.HtmlEncode("?$expand=Attachment,PartnerCaseReference,SlaItem,Kpi")));
            // Console.WriteLine(APIurl);
            string api = APIurl.ToString();
            string ApiJson = await GetHttpContentWithTokenAsync(api, sd.AccessToken);

            return ApiJson;
        }
    }
}
