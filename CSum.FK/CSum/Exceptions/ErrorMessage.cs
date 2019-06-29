using System;
using System.Web;
using CSum.Util;

namespace CSum.Exceptions
{
    /// <summary>
    ///     异常错误信息
    /// </summary>
    [Serializable]
    public class ExceptionMessage
    {
        public ExceptionMessage()
        {
        }

        /// <summary>
        ///     构造函数
        ///     默认显示异常页面
        /// </summary>
        /// <param name="ex">异常对象</param>
        public ExceptionMessage(Exception ex)
            : this(ex, true)
        {
        }

        /// <summary>
        ///     构造函数
        /// </summary>
        /// <param name="ex">异常对象</param>
        /// <param name="isShowException">是否显示异常页面</param>
        public ExceptionMessage(Exception ex, bool isShowException)
        {
            MsgType = ex.GetType().Name;
            Message = ex.InnerException != null ? ex.InnerException.Message : ex.Message;
            StackTrace = ex.StackTrace.Length > 300 ? ex.StackTrace.Substring(0, 300) : ex.StackTrace;
            Source = ex.Source;
            Time = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
            Assembly = ex.TargetSite.Module.Assembly.FullName;
            Method = ex.TargetSite.Name;

            ShowException = isShowException;
            var request = HttpContext.Current.Request;
            IP = Net.Ip;
            UserAgent = request.UserAgent;
            Path = request.Path;
            HttpMethod = request.HttpMethod;
        }

        /// <summary>
        ///     消息类型
        /// </summary>
        public string MsgType { get; set; }

        /// <summary>
        ///     消息内容
        /// </summary>
        public string Message { get; set; }

        /// <summary>
        ///     请求路径
        /// </summary>
        public string Path { get; set; }

        /// <summary>
        ///     程序集名称
        /// </summary>
        public string Assembly { get; set; }

        /// <summary>
        ///     异常参数
        /// </summary>
        public string ActionArguments { get; set; }

        /// <summary>
        ///     请求类型
        /// </summary>
        public string HttpMethod { get; set; }

        /// <summary>
        ///     异常堆栈
        /// </summary>
        public string StackTrace { get; set; }

        /// <summary>
        ///     异常源
        /// </summary>
        public string Source { get; set; }

        /// <summary>
        ///     服务器IP 端口
        /// </summary>
        public string IP { get; set; }

        /// <summary>
        ///     客户端浏览器标识
        /// </summary>
        public string UserAgent { get; set; }


        /// <summary>
        ///     是否显示异常界面
        /// </summary>
        public bool ShowException { get; set; }

        /// <summary>
        ///     异常发生时间
        /// </summary>
        public string Time { get; set; }

        /// <summary>
        ///     异常发生方法
        /// </summary>
        public string Method { get; set; }
    }
}