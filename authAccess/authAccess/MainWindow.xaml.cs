using Microsoft.Identity.Client;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Net.Http;
using System.Windows.Interop;
using Nancy.Helpers;
using static authAccess.ContentHelper;
using System.Threading;
using NPOI.HSSF.UserModel;
using System.IO;

namespace authAccess
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        //https://api.support.microsoft.com/v1/cases/120080525000062/history
        //https://api.support.microsoft.com/v2/cases/120070423000112?$expand=Attachment,PartnerCaseReference,SlaItem,Kpi
       
        SDAuthLib sd = new SDAuthLib();
        //string APIurl = "https://api.support.microsoft.com/v2/cases/120072923002610?$expand=Attachment,PartnerCaseReference,SlaItem,Kpi";
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
        private  void  CallGraphButton_Click(object sender, RoutedEventArgs e)
       {
            ResultWrite();
          //  CaseIdReader.ExcelTest();
        }

        public async void ResultWrite()
        {
            TokenInfoText.Text = string.Empty;
            // SDAuthLib sd = new SDAuthLib();
            //TokenInfoText.Text += $"Token: {sd.GetSDToken()}" + Environment.NewLine;
            sd.GetSDToken();
            // TokenInfoText.Text += $"Token: {sd.AccessToken}" + Environment.NewLine;
            List<String> caseId = CaseIdReader.CaseNumberReader();
            List<Content> con = new List<Content>();
            int j = 0;
            string token = sd.AccessToken;
            foreach (string i in caseId)
            {
                string url = "https://api.support.microsoft.com/v2/cases/";
                StringBuilder APIurl = new StringBuilder(url);
                APIurl.Append(String.Format("{0}", HttpUtility.HtmlEncode(i)));
                APIurl.Append(String.Format("{0}", HttpUtility.HtmlEncode("?$expand=Attachment,PartnerCaseReference,SlaItem,Kpi")));
                // Console.WriteLine(APIurl);
                string api = APIurl.ToString();
                string ReadApi = await GetHttpContentWithTokenAsync(api, sd.AccessToken);
                TokenInfoText.Text += $"Token: {sd.AccessToken}" + Environment.NewLine;
                TokenInfoText.Text += $"JSON: {ReadApi}" + Environment.NewLine;
                dynamic dobj = JsonConvert.DeserializeObject<dynamic>(ReadApi);
                Content data = new Content()
                {
                    Number = j + 1,
                    CaseNumber = caseId[j],
                    Name = dobj["AgentId"].ToString(),
                    Topic = dobj["Title"].ToString(),
                    SupportCountry = dobj["SupportCountry"].ToString(),
                    ServiceLevel = dobj["EntitlementInformation"]["ServiceLevel"].ToString(),
                };
                con.Add(data);

                ResultText.Text += $"CaseNumber: {data.CaseNumber}" + Environment.NewLine;
                ResultText.Text += $"AgentId: {data.Name}" + Environment.NewLine;
                ResultText.Text += $"SupportCountry: {data.SupportCountry}" + Environment.NewLine;
                ResultText.Text += $"ServiceLevel: {data.ServiceLevel}" + Environment.NewLine;
                //ResultText.Text += $"Severity: {data.}" + Environment.NewLine;
                ResultText.Text += $"Topic: {data.Topic}" + Environment.NewLine;
                APIurl.Clear();
                j++;
            }
            //导出：将数据库中的数据，存储到一个excel中
            //1、查询数据库数据  
            //2、  生成excel
            //2_1、生成workbook
            //2_2、生成sheet
            //2_3、遍历集合，生成行
            //2_4、根据对象生成单元格
            HSSFWorkbook workbook = new HSSFWorkbook();
            //创建工作表
            var sheet = workbook.CreateSheet("信息表");
            //创建标题行（重点） 从0行开始写入
            var row = sheet.CreateRow(0);
            //创建单元格
            var cellid = row.CreateCell(0);
            cellid.SetCellValue("nums");
           
            var cellname = row.CreateCell(1);
            cellname.SetCellValue("CaseNumber");
            
            var cellpwd = row.CreateCell(2);
            cellpwd.SetCellValue("Name");
            var cellTopic = row.CreateCell(3);
            cellTopic.SetCellValue("Topic");
            var cellReg = row.CreateCell(4);
            cellReg.SetCellValue("Region");
            var cellsup = row.CreateCell(5);
            cellsup.SetCellValue("Support Level");
            //遍历集合，生成行
            int index = 1; //从1行开始写入
            for (int i = 0; i < caseId.Count; i++)
            {
                int x = index + i;
                var rowi = sheet.CreateRow(x);
                var ids = rowi.CreateCell(0);
                ids.SetCellValue(con[i].Number);
                var num = rowi.CreateCell(1);
                num.SetCellValue(con[i].CaseNumber);
                var name = rowi.CreateCell(2);
                name.SetCellValue(con[i].Name);
                var topic = rowi.CreateCell(3);
                topic.SetCellValue(con[i].Topic);
                var region = rowi.CreateCell(4);
                region.SetCellValue(con[i].SupportCountry);
                var supportLevel = rowi.CreateCell(5);
                supportLevel.SetCellValue(con[i].ServiceLevel);
                
            }
            for(int k =0; k<6;k++)
            {
                sheet.AutoSizeColumn(k);
            }
            
            //DirectoryInfo di = new DirectoryInfo(@"C:\Users\Daisy\Desktop\inf.xls");
            String rootFolder = @"C:\Users\Daisy\Desktop";
            string file = "inf.xls";
            try
            {
                if (File.Exists(Path.Combine(rootFolder, file)))
                {
                   File.Delete(Path.Combine(rootFolder, file));
                }

            }
            catch (Exception e)
            {
                Console.WriteLine("This process failed:{0}", e.Message);
            }
            string w = @"C:\Users\Daisy\Desktop\" + "CaseReader_" + DateTime.Now.ToString("yyyy-MM-dd-HH-mm-ss") + ".xls";
            FileStream file1 = new FileStream(w, FileMode.CreateNew, FileAccess.Write);
            workbook.Write(file1);
            file1.Dispose();
        }
        private  void SignOutButton_Click(object sender, RoutedEventArgs e)
        {
            if (sd.GetSDToken())
            {
                try
                {
                   
                }
                catch (MsalException ex)
                {
                    ResultText.Text = $"Error signing-out user: {ex.Message}";
                }
            }
        }

        private void TokenInfoText_TextChanged(object sender, TextChangedEventArgs e)
        {

        }
    }
}
