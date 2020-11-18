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
        private  void CallGraphButton_Click(object sender, RoutedEventArgs e)
        {
            
            SDAuthLib sd = new SDAuthLib();
            sd.GetSDToken();
            //SDAuthLib.ReadOrCreateADALTokenCache();
           
            string token = sd.AccessToken;
            ResultWrite(token);
            sd.ClearToken();
                //  CaseIdReader.ExcelTest();     
        }
        public async void ResultWrite(string token)
        {
            CaseIdReader cr = new CaseIdReader();
            List<Content> CR = await cr.CaseNumberReaderAsync(token);
            //sort the List ,from the old case to the newest case
            var sortedData =
               (from s in CR
                select new
                {
                    s.caseId,
                    s.time,
                    s.alias,
                    s.severity,
                    s.ServiceLevel,
                    s.isTask,
                    s.Topic,
                    s.SupportCountry,
                    s.iTime,
                    s.iDate,
                    s.SLA,
                    s.vertical
                }).OrderBy(x => x.caseId).ToList();
           
            foreach (var i in sortedData)
            {
                ResultText.Text += "CaseId:"+"     "+i.caseId + Environment.NewLine;
                
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
            cellid.SetCellValue("CaseId");

            var date = row.CreateCell(1);
            date.SetCellValue("Date");
            var cellpwd = row.CreateCell(2);
            cellpwd.SetCellValue("Time");
            var time = row.CreateCell(3);
            time.SetCellValue("Name");
            var cellTopic = row.CreateCell(4);
            cellTopic.SetCellValue("Topic");
            var cellReg = row.CreateCell(5);
            cellReg.SetCellValue("Region");
            var cellsup = row.CreateCell(6);
            cellsup.SetCellValue("Support Level");
            
            var sev = row.CreateCell(7);
            sev.SetCellValue("Severity");
            var sla = row.CreateCell(8);
            sla.SetCellValue("  SLA  ");
            var vertical = row.CreateCell(9);
            vertical.SetCellValue("Vertical");
            var it = row.CreateCell(10);
            it.SetCellValue("Item Type");

            //遍历集合，生成行
            int index = 1; //从1行开始写入
            for (int i = 0; i < sortedData.Count; i++)
            {
                int x = index + i;
                var rowi = sheet.CreateRow(x);
                var ids = rowi.CreateCell(0);
                ids.SetCellValue(sortedData[i].caseId);
                var d = rowi.CreateCell(1);
                d.SetCellValue(sortedData[i].iDate);
                var t = rowi.CreateCell(2);
                t.SetCellValue(sortedData[i].iTime);
                var name = rowi.CreateCell(3);
                name.SetCellValue(sortedData[i].alias);
                var topic = rowi.CreateCell(4);
                topic.SetCellValue(sortedData[i].Topic);
                var region = rowi.CreateCell(5);
                region.SetCellValue(sortedData[i].SupportCountry);
                var supportLevel = rowi.CreateCell(6);
                supportLevel.SetCellValue(sortedData[i].ServiceLevel);
                var Severity = rowi.CreateCell(7);
                Severity.SetCellValue(sortedData[i].severity);
                var S = rowi.CreateCell(8);
                S.SetCellValue(sortedData[i].SLA);
                var V = rowi.CreateCell(9);
                V.SetCellValue(sortedData[i].vertical);

                var ItemType = rowi.CreateCell(10);
                ItemType.SetCellValue(sortedData[i].isTask);

            }
            for (int k = 0; k < 14; k++)
            {
                if(k !=4)
                sheet.AutoSizeColumn(k);
            }

            //DirectoryInfo di = new DirectoryInfo(@"C:\Users\Daisy\Desktop\inf.xls");
            //string path = Path.Combine(AppDomain.CurrentDomain.BaseDirectory);
            //string path = "E://holiday.json";
            // path = Path.Combine(path + "\\.xls");
            string dir = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory);
            
            string w = dir + "\\CaseReader_" + DateTime.Now.ToString("yyyy-MM-dd-HH-mm-ss") + ".xls";
            ResultText.Text += w + Environment.NewLine;
            FileStream file1 = new FileStream(w, FileMode.CreateNew, FileAccess.Write);
            workbook.Write(file1);
            file1.Dispose();
            TokenInfoText.Text += "Data has been written into excel" + Environment.NewLine;
            /*SDAuthLib sd = new SDAuthLib();
           
            var e  = sd.Expiry;
            var a = sd.BypassTokenCache;
            var c = sd.GetSDToken();
            TokenInfoText.Text += "sd.Expiry                      "+e + Environment.NewLine;
            TokenInfoText.Text += "sd.BypassTokenCache     " + a + Environment.NewLine;
            TokenInfoText.Text += " sd.GetSDToken()         " + c + Environment.NewLine;
            */
        }
    
        //https://api.support.microsoft.com/v1/cases/120080525000062/history
        //https://api.support.microsoft.com/v2/cases/120070423000112?$expand=Attachment,PartnerCaseReference,SlaItem,Kpi

        //SDAuthLib sd = new SDAuthLib();
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

        private void SignOutButton_Click(object sender, RoutedEventArgs e)
        {

        }


    }
}
