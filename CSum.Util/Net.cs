using System.Net;
using System.Net.Sockets;
using System.Web;

namespace CSum.Util
{
    /// <summary>
    ///     网络操作
    /// </summary>
    public class Net
    {
        #region Browser(获取浏览器信息)

        /// <summary>
        ///     获取浏览器信息
        /// </summary>
        public static string Browser
        {
            get
            {
                if (HttpContext.Current == null)
                    return string.Empty;
                var browser = HttpContext.Current.Request.Browser;
                return string.Format("{0} {1}", browser.Browser, browser.Version);
            }
        }

        #endregion

        #region Ip(获取Ip)

        /// <summary>
        ///     获取Ip
        /// </summary>
        public static string Ip
        {
            get
            {
                var result = string.Empty;
                if (HttpContext.Current != null)
                    result = GetWebClientIp();
                if (result.IsEmpty())
                    result = GetLanIp();
                return result;
            }
        }

        /// <summary>
        ///     获取Web客户端的Ip
        /// </summary>
        private static string GetWebClientIp()
        {
            var ip = GetWebRemoteIp();
            foreach (var hostAddress in Dns.GetHostAddresses(ip))
                if (hostAddress.AddressFamily == AddressFamily.InterNetwork)
                    return hostAddress.ToString();
            return string.Empty;
        }

        /// <summary>
        ///     获取Web远程Ip
        /// </summary>
        private static string GetWebRemoteIp()
        {
            return HttpContext.Current.Request.ServerVariables["HTTP_X_FORWARDED_FOR"] ??
                   HttpContext.Current.Request.ServerVariables["REMOTE_ADDR"];
        }

        /// <summary>
        ///     获取局域网IP
        /// </summary>
        private static string GetLanIp()
        {
            foreach (var hostAddress in Dns.GetHostAddresses(Dns.GetHostName()))
                if (hostAddress.AddressFamily == AddressFamily.InterNetwork)
                    return hostAddress.ToString();
            return string.Empty;
        }

        #endregion

        #region Host(获取主机名)

        /// <summary>
        ///     获取主机名
        /// </summary>
        public static string Host => HttpContext.Current == null ? Dns.GetHostName() : GetWebClientHostName();

        /// <summary>
        ///     获取Web客户端主机名
        /// </summary>
        private static string GetWebClientHostName()
        {
            if (!HttpContext.Current.Request.IsLocal)
                return string.Empty;
            var ip = GetWebRemoteIp();
            var result = Dns.GetHostEntry(IPAddress.Parse(ip)).HostName;
            if (result == "localhost.localdomain")
                result = Dns.GetHostName();
            return result;
        }

        /// <summary>
        ///     获取当前部署站点的协议名 + 域名 + 端口号
        /// </summary>
        public static string Url
        {
            get
            {
                var context = HttpContext.Current;
                if (context == null) return "";

                var t = context.Request.Url.AbsoluteUri.IndexOf(context.Request.Url.Authority);
                return context.Request.Url.AbsoluteUri.Substring(0, t) + context.Request.Url.Authority;
            }
        }

        #endregion
    }
}