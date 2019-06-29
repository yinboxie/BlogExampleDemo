using System.Security.Cryptography;
using System.Text;

namespace CSum.Util
{
    /// <summary>
    ///     MD5加密帮助类
    /// </summary>
    public class Md5Helper
    {
        /// <summary>
        ///     MD5加密 默认32位大写
        /// </summary>
        /// <param name="str">加密字符</param>
        /// <returns></returns>
        public static string MD5(string str)
        {
            return MD5(str, 32);
        }

        /// <summary>
        ///     MD5加密 默认大写
        /// </summary>
        /// <param name="str">加密字符</param>
        /// <param name="code">位数 16或32</param>
        /// <returns></returns>
        public static string MD5(string str, int code)
        {
            return MD5(str, code, true);
        }

        /// <summary>
        ///     MD5加密
        /// </summary>
        /// <param name="str">加密字符</param>
        /// <param name="code">位数 16或32</param>
        /// <param name="isUp">大小写 true 大写 false 小写</param>
        /// <returns></returns>
        public static string MD5(string str, int code, bool isUp)
        {
            if (isUp)
                return MD5Up(str, code);
            return MD5Lower(str, code);
        }

        #region "大写 MD5加密"

        /// <summary>
        ///     大写 MD5加密
        /// </summary>
        /// <param name="str">加密字符</param>
        /// <param name="code">位数,默认32位</param>
        /// <returns></returns>
        public static string MD5Up(string str, int code = 32)
        {
            if (string.IsNullOrEmpty(str))
                return str;
            var sb = new StringBuilder(code);
            var md5 = System.Security.Cryptography.MD5.Create();
            var output = md5.ComputeHash(Encoding.UTF8.GetBytes(str));
            for (var i = 0; i < output.Length; i++)
                sb.Append(output[i].ToString("X").PadLeft(2, '0'));
            return sb.ToString();
        }

        #endregion

        #region "小写 MD5加密"

        /// <summary>
        ///     小写 MD5加密
        /// </summary>
        /// <param name="str">加密字符</param>
        /// <param name="code">位数,默认32位</param>
        /// <returns></returns>
        public static string MD5Lower(string str, int code = 32)
        {
            //32位小写MD5加密
            MD5 md5 = new MD5CryptoServiceProvider();
            var data = Encoding.UTF8.GetBytes(str);
            var md5data = md5.ComputeHash(data);
            var sb = new StringBuilder(code);
            for (var i = 0; i < md5data.Length; i++) sb.Append(md5data[i].ToString("x").PadLeft(2, '0'));
            return sb.ToString();
        }

        #endregion
    }
}