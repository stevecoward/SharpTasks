using System.IO;
using System.Net;
using System.Security.Authentication;

namespace SharpTasks.Reconnaissance
{
    /// <summary>
    ///     Contains methods and properties for anything related to IP addresses.
    /// </summary>
    public class IpAddress
    {
        private const SslProtocols _Tls12 = (SslProtocols) 0x00000C00;
        private const SecurityProtocolType Tls12 = (SecurityProtocolType) _Tls12;

        /// <summary>
        ///     Fetches the host's external IP address by making a web request to ifconfig.co.
        /// </summary>
        /// <returns>Host's external IP address.</returns>
        public static string GetExternalIp()
        {
            const string url = "https://ifconfig.co/ip";
            var ip = "";

            // Setting TLS v1.2 as the protocol as ifconfig.co won't accept other versions.
            ServicePointManager.SecurityProtocol = Tls12;
            var request = WebRequest.Create(url);
            var response = request.GetResponse();
            using (var dataStream = response.GetResponseStream())
            {
                var reader = new StreamReader(dataStream);
                ip = reader.ReadToEnd();
            }

            response.Close();

            return ip.Replace("\n", "");
        }
    }
}