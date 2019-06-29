using System;

namespace CSum.Logging
{
    /// <summary>
    ///     日志记录接口
    /// </summary>
    public interface ILogger
    {
        #region 级别日志

        /// <summary>
        ///     调试期间的日志
        /// </summary>
        /// <param name="message"></param>
        void Debug(string message);

        /// <summary>
        ///     将message记录到日志文件
        /// </summary>
        /// <param name="message"></param>
        void Info(string message);

        /// <summary>
        ///     引起警告的日志
        /// </summary>
        /// <param name="message"></param>
        void Warn(string message);

        /// <summary>
        ///     异常发生的日志
        /// </summary>
        /// <param name="message"></param>
        void Error(Exception ex);

        /// <summary>
        ///     引起程序终止的日志
        /// </summary>
        /// <param name="message"></param>
        void Fatal(string message);

        #endregion
    }
}