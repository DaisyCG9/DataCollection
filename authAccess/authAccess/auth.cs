using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Win32;
using System.IO;
using System.Net;
using System.Security.Cryptography;
using System.Web;
//using MimeKit.Encodings;
using Microsoft.IdentityModel.Clients.ActiveDirectory;

namespace authAccess
{

    class SDAuthLib
    {

        public const string logonUrl = "https://login.microsoftonline.com/common/";
        private const string serviceName = "MSaaS";
        private object serializer = new object();
        private const string regPath = "Software\\Microsoft\\MSDTools";
        private const string adalClientId = "ad9a38dc-2fa5-4863-9557-4f9b4a23e44b";
        private const string adalResourceId = "https://api.support.microsoft.com";
        private const string adalRedirectUrl = "https://casebuddy.microsoft.com"; 
        private static bool bTokenCacheDeserializeAttempted = false;
        private DateTime expiry = DateTime.MinValue;

        private static TokenCache myTokenCache = null;
        //  private string cookieAuthUrl = "";
        public string extraParams;
        private string userUPN = "@microsoft.com";
        private string token;
        public bool BypassTokenCache
        {
            get
            {
                return bBypassTokenCache;
            }
            set
            {
                bBypassTokenCache = value;
            }
        }
        public DateTime Expiry => expiry;


        private bool bBypassTokenCache;
        //	private bool bNoImplicit = true;
        public bool bSilent = true;
        private bool bAuthFailed = false;


        public string AccessToken
        {
            get { return token; }
            set { token = value; }
        }

        public SDAuthLib()
        {
            ReadOrCreateADALTokenCache();
        }
        public static void ReadOrCreateADALTokenCache()
        {   //bTokenCacheDeserializeAttempted初始化为false
            if (!bTokenCacheDeserializeAttempted)
            {
                RegistryKey registryKey = null;
                myTokenCache = null;
                try
                {
                    //CreateSubKey(String, RegistryKeyPermissionCheck)
                    //Creates a new subkey or opens an existing subkey for write access, using the specified permission check option.
                    //Software\\Microsoft\\MSDTools\\OAuthCache  --The name or path of the subkey to create or open. This string is not case-sensitive.
                    //RegistryKeyPermissionCheck
                    //One of the enumeration values that specifies whether the key is opened for read or read / write access.
                    registryKey = Registry.CurrentUser.CreateSubKey("Software\\Microsoft\\MSDTools\\OAuthCache", RegistryKeyPermissionCheck.ReadSubTree);
                    byte[] array = (byte[])registryKey.GetValue("ADALTokenCache", new byte[1]);
                    if (array.Length > 1)
                    {
                        myTokenCache = new TokenCache(array);
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine("Cannot read/use ADAL token cache: " + ex.Message + " ReadOrCreateADALTokenCache");
                    myTokenCache = null;
                }
                registryKey?.Close();
                bTokenCacheDeserializeAttempted = true;
                if (myTokenCache == null)
                {
                    myTokenCache = new TokenCache();
                }
            }
        }

        public bool GetSDToken(bool bSilent = true)
        {
            if (bBypassTokenCache)
            {
                ClearToken();
            }

            ReadSDToken(serviceName, ref token, ref expiry);

            if (!Valid())
            {
                lock (serializer)
                {
                    AuthenticationContext ac = new AuthenticationContext(string.IsNullOrEmpty(logonUrl) ? "https://login.microsoftonline.com/72f988bf-86f1-41af-91ab-2d7cd011db47/" : logonUrl, myTokenCache);

                    try
                    {

                        Task<AuthenticationResult> task = ac.AcquireTokenSilentAsync(adalResourceId, adalClientId, new UserIdentifier(userUPN, UserIdentifierType.OptionalDisplayableId));
                        task.Wait(30000);
                        token = task.Result.AccessToken;
                        Console.WriteLine("Token is " + token);
                        ExtractExpiry();
                        SaveADALTokenCache();
                        WriteSDToken(serviceName, token, expiry);
                        return true;

                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine("Authentication failed with error: " + ex.Message.ToString());
                        bAuthFailed = true;

                    }

                    if (bAuthFailed && bSilent)
                    {
                        try
                        {
                            //OVERLOADS   PlatformParameters(PromptBehavior)

                            Task<AuthenticationResult> task2 = ac.AcquireTokenAsync(adalResourceId, adalClientId,
                             new Uri(adalRedirectUrl), new PlatformParameters(PromptBehavior.Always),
                             new UserIdentifier(userUPN, UserIdentifierType.OptionalDisplayableId), extraParams);
                            task2.Wait(30000);
                            token = task2.Result.AccessToken;
                            Console.WriteLine("Token is " + token);
                            ExtractExpiry();
                            SaveADALTokenCache();
                            WriteSDToken(serviceName, token, expiry);
                            return true;
                        }
                        catch (Exception ex)
                        {

                            Console.WriteLine("Authentication failed with error: " + ex.Message.ToString());
                            bAuthFailed = true;

                        }
                    }

                    if (!bAuthFailed)
                    {
                        return true;
                    }

                }

            }
            return Valid();
        }

        public static string MyDecodeBase64(string inp)
        {
            try
            {
                //Base64Decoder方法已经不可用
                /*Base64Decoder base64Decoder = new Base64Decoder(inp.ToCharArray());
               return Encoding.UTF8.GetString(base64Decoder.GetDecoded());*/
                //  https://stackoverflow.com/questions/11743160/how-do-i-encode-and-decode-a-base64-string
                var base64EncodedBytes = System.Convert.FromBase64String(inp);
                return System.Text.Encoding.UTF8.GetString(base64EncodedBytes);

            }
            catch (Exception)
            {
                return "INVALID";
            }
        }

        public void ClearToken()
        {
            expiry = DateTime.MinValue;
            token = "";
            //         bool bCookieAuth2 = bCookieAuth;
        }

        public bool Valid()
        {
            if (!string.IsNullOrEmpty(token))
            {
                return expiry > DateTime.Now;
            }
            return false;
        }

        private void ExtractExpiry()
        {
            if (string.IsNullOrEmpty(token))
            {
                expiry = DateTime.MinValue;
            }
            else
            {
                try
                {
                    string[] array = token.Split(".".ToCharArray());
                    if (array.Length == 3)
                    {
                        expiry = DateTime.Now.AddMinutes(58.0);
                    }
                    string text = MyDecodeBase64(array[1]);
                    int num = text.IndexOf("\"exp\":");
                    if (num > 0)
                    {
                        num += 6;
                        int num2 = text.IndexOf(",", num);
                        if (num2 > num)
                        {
                            int result = 0;
                            if (int.TryParse(text.Substring(num, num2 - num), out result))
                            {
                                expiry = new DateTime(1970, 1, 1, 0, 0, 0).AddSeconds(result);
                                expiry = expiry.ToLocalTime().AddMinutes(-2.0);
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    expiry = DateTime.Now.AddMinutes(59.0);
                    
                }
            }
        }

        private static void SaveADALTokenCache()
        {
            if (myTokenCache != null)
            {
                try
                {
                    Registry.CurrentUser.CreateSubKey("Software\\Microsoft\\MSDTools\\OAuthCache", RegistryKeyPermissionCheck.ReadWriteSubTree).SetValue("ADALTokenCache", myTokenCache.Serialize());
                }
                catch (Exception ex)
                {
                    Console.WriteLine("Cannot store ADAL token cache: " + ex.Message + " SaveADALTokenCache");
                }
            }
        }

        public static void WriteSDToken(string svc, string token, DateTime expiry)
        {
            if (!(DateTime.Now >= expiry))
            {
                RegistryKey registryKey = null;
                try
                {
                    registryKey = Registry.CurrentUser.CreateSubKey("Software\\Microsoft\\MSDTools\\OAuthCache", RegistryKeyPermissionCheck.ReadWriteSubTree);
                    registryKey.SetValue(svc + "-Token-Expiry", expiry.ToBinary());
                    registryKey.SetValue(svc + "-Token", token);
                    //               registryKey.SetValue(svc + "-BrowserAuthed", bBrowserAuth ? 1 : 0);
                }
                catch (Exception)
                {
                }
                registryKey?.Close();
            }
        }

        //ref --to pass an argument to a method by reference
        public static bool ReadSDToken(string svc, ref string token, ref DateTime expiry)
        {
            RegistryKey registryKey = null;
            expiry = DateTime.MinValue;
            token = "";
            try
            {
                registryKey = Registry.CurrentUser.CreateSubKey("Software\\Microsoft\\MSDTools\\OAuthCache", RegistryKeyPermissionCheck.ReadSubTree);
                if (long.TryParse((string)registryKey.GetValue(svc + "-Token-Expiry", long.MinValue.ToString()), out long result))
                {
                    expiry = DateTime.FromBinary(result);
                }
                token = (string)registryKey.GetValue(svc + "-Token", "");
            }
            catch (Exception)
            {
            }
            registryKey?.Close();
            return expiry > DateTime.Now.AddMinutes(-1.0);
        }


    }
}
