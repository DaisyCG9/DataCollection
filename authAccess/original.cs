using Microsoft.Identity.Client;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Interop;
namespace authAccess
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        string[] scopes = new string[] { "user.read" };
        private void TokenInfoText_TextChanged(object sender, TextChangedEventArgs e)
        {

        }
        public async Task<string> GetHttpContentWithTokenAsync(string url, string token)
        {
            var httpClient = new System.Net.Http.HttpClient();
            System.Net.Http.HttpResponseMessage response;
            try
            {
                var request = new System.Net.Http.HttpRequestMessage(System.Net.Http.HttpMethod.Get, url);
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

        /*private async void  CallGraphButton_ClickAsync(object sender, RoutedEventArgs e)
        {
            SDAuthLib sd = new SDAuthLib();
           // string url = "https://api.support.microsoft.com";
            //TokenInfoText.Text += "Token: s" + Environment.NewLine;
            TokenInfoText.Text += $"Token: {sd.GetSDToken()}" + Environment.NewLine;
            TokenInfoText.Text += $"Token: {sd.AccessToken}" + Environment.NewLine;
            //TokenInfoText.Text += $"Token: {await GetHttpContentWithTokenAsync(url, sd.AccessToken)}" + Environment.NewLine;

        }*/
        private void CallGraphButton_Click(object sender, RoutedEventArgs e)
        {
            TokenInfoText.Text = string.Empty;
            SDAuthLib sd = new SDAuthLib();
            //TokenInfoText.Text += $"Token: {sd.GetSDToken()}" + Environment.NewLine;
            sd.GetSDToken(true);
            TokenInfoText.Text += $"Token: {sd.AccessToken}" + Environment.NewLine;


        }
    }
}
