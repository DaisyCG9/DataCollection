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

namespace authAccess
{
    class CaseIdReader
    {
        public static List<String> CaseNumberReader()
        {
            List<string> id = new List<string>();
            var mails = OutlookEmails.ReadMailItems();

            foreach (var mail in mails)
            {
                String pattern1 = @"[1]\d{14}";
                String CaseReader = mail.EmailSubject;
                Match match1 = Regex.Match(CaseReader, pattern1, RegexOptions.IgnoreCase);
                if (match1.Success)
                {
                    id.Add(match1.Value);
                }    
            }
            //Sort the numbers from the oldest to the lastest.
            id.Sort();
            //remove the dulplicate numbers
            List<string> id1=id.Distinct().ToList();
            //Console.WriteLine(string.Join("\n", id1));
           
            return id1;
        }
        
    }
}
