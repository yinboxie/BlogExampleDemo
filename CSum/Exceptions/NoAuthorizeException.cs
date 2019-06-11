using System;

namespace CSum.Exceptions
{
    /// <summary>
    ///     没有被授权的异常
    /// </summary>
    public class NoAuthorizeException : Exception
    {
        public NoAuthorizeException(string message)
            : base(message)
        {
        }
    }
}