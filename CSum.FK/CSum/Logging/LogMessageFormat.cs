using System;
using System.Text;
using CSum.Util;

namespace CSum.Logging
{
    /// <summary>
    ///     日志格式器
    /// </summary>
    public static class LogMessageFormat
    {
        /// <summary>
        ///     生成警告
        /// </summary>
        /// <param name="content">警告内容</param>
        /// <returns></returns>
        public static string WarnFormat(string content)
        {
            // var user = SessionManager.Instance.Current();
            dynamic user = null;
            var strInfo = new StringBuilder();
            strInfo.Append("1. 警告: >> 操作时间: " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "   操作人: " +
                           (user == null ? "" : user.UserId) + " \r\n");
            strInfo.Append("2. 主机: " + Net.Host + "   Ip  : " + Net.Ip + "   浏览器: " + Net.Browser + "    \r\n");
            strInfo.Append("3. 内容: " + content + "\r\n");
            strInfo.Append(
                "-----------------------------------------------------------------------------------------------------------------------------\r\n");
            return strInfo.ToString();
        }

        /// <summary>
        ///     生成错误信息,保存至文件的文本
        /// </summary>
        /// <param name="ex">异常对象</param>
        /// <returns></returns>
        public static string ErrorFormat(Exception ex)
        {
            //var user = SessionManager.Instance.Current();
            dynamic user = null;
            var strInfo = new StringBuilder();
            strInfo.Append("1. 错误: >> 操作时间: " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "   操作人: " +
                           (user == null ? "" : user.UserId) + " \r\n");
            strInfo.Append("2. 程序集: " + ex.TargetSite.Module.Assembly.FullName + "    \r\n");
            strInfo.Append("3. 异常方法: " + ex.TargetSite.Name + " \r\n");
            strInfo.Append("4. 主机: " + Net.Host + "   Ip  : " + Net.Ip + "   浏览器: " + Net.Browser + "    \r\n");
            strInfo.Append("5. 异常: " + (ex.InnerException != null ? ex.InnerException.Message : ex.Message) + "\r\n");
            strInfo.Append("6. 调用堆栈: " + ex.StackTrace + "\r\n");
            strInfo.Append(
                "-----------------------------------------------------------------------------------------------------------------------------\r\n");
            return strInfo.ToString();
        }

        /// <summary>
        ///     生成异常信息,保存至数据库的文本
        /// </summary>
        /// <param name="ex">异常对象</param>
        /// <returns></returns>
        public static string ExceptionFormat(Exception ex)
        {
            //var user = SessionManager.Instance.Current();
            dynamic user = null;
            var strInfo = new StringBuilder();
            strInfo.Append("1. 错误: >> 操作时间: " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "   操作人: " +
                           (user == null ? "" : user.UserId) + " \r\n");
            strInfo.Append("2. 程序集: " + ex.TargetSite.Module.Assembly.FullName + "    \r\n");
            strInfo.Append("3. 异常方法: " + ex.TargetSite.Name + " \r\n");
            strInfo.Append("4. 主机: " + Net.Host + "   Ip  : " + Net.Ip + "   浏览器: " + Net.Browser + "    \r\n");
            strInfo.Append("5. 异常: " + (ex.InnerException != null ? ex.InnerException.Message : ex.Message) + "\r\n");
            strInfo.Append(
                "-----------------------------------------------------------------------------------------------------------------------------\r\n");
            return strInfo.ToString();
        }
    }
}