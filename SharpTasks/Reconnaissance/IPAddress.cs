using System;
using System.IO;
using System.Net;
using System.Security.Authentication;

namespace SharpTasks.Reconnaissance
{
    public class IPAddress
    {
        const SslProtocols _Tls12 = (SslProtocols)0x00000C00;
        const SecurityProtocolType Tls12 = (SecurityProtocolType)_Tls12;

        public static string GetExternalIP()
        {
            const string url = "https://ifconfig.co/ip";
            string ip = "";

            ServicePointManager.SecurityProtocol = Tls12;
            WebRequest request = WebRequest.Create(url);
            WebResponse response = request.GetResponse();
            using (Stream dataStream = response.GetResponseStream())
            {
                StreamReader reader = new StreamReader(dataStream);
                ip = reader.ReadToEnd();
            }

            response.Close();

            return ip.Replace("\n", "");
        }
    }
}