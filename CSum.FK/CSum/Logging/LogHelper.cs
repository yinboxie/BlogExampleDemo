using System;
using Common.Logging;

namespace CSum.Logging
{
    /// <summary>
    ///     日志记录帮助类
    /// </summary>
    public static class LogHelper
    {
        #region 级别日志

        /// <summary>
        ///     调试期间的日志
        /// </summary>
        /// <param name="message"></param>
        public static void Debug(string message)
        {
            LogManager.GetLogger("DebugLog").Debug(message);
        }

        /// <summary>
        ///     跟踪日志
        /// </summary>
        /// <param name="message"></param>
        public static void Trace(string message)
        {
            LogManager.GetLogger("TraceLog").Info(message);
        }

        /// <summary>
        ///     将message记录到日志文件
        /// </summary>
        /// <param name="message"></param>
        public static void Info(string message)
        {
            LogManager.GetLogger("InfoLog").Info(message);
        }

        /// <summary>
        ///     引起警告的日志
        /// </summary>
        /// <param name="message"></param>
        public static void Warn(string message)
        {
            LogManager.GetLogger("WarnLog").Warn(LogMessageFormat.WarnFormat(message));
        }

        /// <summary>
        ///     异常发生的日志
        /// </summary>
        /// <param name="ex"></param>
        public static void Error(Exception ex)
        {
            LogManager.GetLogger("ErrorLog").Error(LogMessageFormat.ErrorFormat(ex));
        }

        /// <summary>
        ///     引起程序终止的日志
        /// </summary>
        /// <param name="message"></param>
        public static void Fatal(string message)
        {
            LogManager.GetLogger("FatalLog").Fatal(message);
        }

        #endregion
    }
}