using System;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Web;
using System.Web.Http.Controllers;
using System.Web.Http.Filters;
using CSum.Web;
using CSum.Exceptions;
using CSum.Util;

namespace CSum.WebApi
{
    /// <summary>
    ///     安全拦截过滤器
    /// </summary>
    public class HandlerSecretAttribute : ActionFilterAttribute
    {
        private readonly ExcuteMode _customMode;

        /// <summary>默认构造</summary>
        /// <param name="Mode">认证模式</param>
        public HandlerSecretAttribute(ExcuteMode Mode)
        {
            _customMode = Mode;
        }

        /// <summary>
        ///     安全校验
        /// </summary>
        /// <param name="filterContext"></param>
        public override void OnActionExecuting(HttpActionContext filterContext)
        {
            //是否忽略权限验证
            if (_customMode == ExcuteMode.Ignore) return;

            //从http请求的头里面获取AppId
            var request = filterContext.Request;
            var method = request.Method.Method;
            var appId = ""; //客户端应用唯一标识
            long timeStamp; //时间戳， 校验10分钟内有效
            var signature = ""; //参数签名，去除空参数,按字母倒序排序进行Md5签名 为了提高传参过程中，防止参数被恶意修改，在请求接口的时候加上sign可以有效防止参数被篡改
            try
            {
                appId = request.Headers.GetValues("appId").SingleOrDefault();
                timeStamp = Convert.ToInt64(request.Headers.GetValues("timeStamp").SingleOrDefault());
                signature = request.Headers.GetValues("signature").SingleOrDefault();
            }
            catch (Exception ex)
            {
                throw new UserFriendlyException("签名参数异常:" + ex.Message);
            }

            #region 安全校验

            //appId是否为可用的
            if (!VerifyAppId(appId)) throw new UserFriendlyException("AppId不被允许访问:" + appId);

            var tonow = Time.StampToDateTime(timeStamp.ToString());

            //请求是否超时
            var expires_minute = tonow.Minute - DateTime.Now.Minute;
            if (expires_minute > 10 || expires_minute < -10) throw new UserFriendlyException("接口请求超时:" + expires_minute);


            //根据请求类型拼接参数
            var form = HttpContext.Current.Request.QueryString;
            var data = string.Empty;
            switch (method)
            {
                case "POST":
                    if (form.Count > 0)
                    {
                        data = GetFormQueryString(form);
                    }
                    else
                    {
                        var stream = HttpContext.Current.Request.InputStream;
                        stream.Position = 0;
                        var responseJson = string.Empty;
                        var streamReader = new StreamReader(stream);
                        data = streamReader.ReadToEnd();
                        stream.Position = 0;
                    }
                    break;
                case "GET":
                    data = GetFormQueryString(form);
                    break;
            }
            var result = Validate(timeStamp.ToString(), data, signature);
            if (!result)
                throw new UserFriendlyException("无效签名");
            base.OnActionExecuting(filterContext);

            #endregion
        }

        #region 私有方法

        /// <summary>
        ///     验证appId是否被允许
        /// </summary>
        /// <param name="appId"></param>
        /// <returns></returns>
        private bool VerifyAppId(string appId)
        {
            if (appId.IsEmpty()) return false;
            return ConfigHelper.GetValue("AllowAppId").IndexOf(appId) > -1;
        }

        /// <summary>
        /// 表单查询参数， url上直接接参数时,通过此方法获取
        /// </summary>
        /// <param name="form"></param>
        /// <returns></returns>
        private string GetFormQueryString(NameValueCollection form)
        {
            //第一步：取出所有get参数
            IDictionary<string, string> parameters = new Dictionary<string, string>();
            for (var f = 0; f < form.Count; f++)
            {
                var key = form.Keys[f];
                parameters.Add(key, form[key]);
            }

            // 第二步：把字典按Key的字母顺序排序
            IDictionary<string, string> sortedParams = new SortedDictionary<string, string>(parameters);
            var dem = sortedParams.GetEnumerator();

            // 第三步：把所有参数名和参数值串在一起
            var query = new StringBuilder();
            while (dem.MoveNext())
            {
                var key = dem.Current.Key;
                var value = dem.Current.Value;
                if (!string.IsNullOrEmpty(key)) query.Append(key).Append(value);
            }

            return query.ToString();
        }

        /// <summary>
        ///     签名校验
        /// </summary>
        /// <param name="timeStamp"></param>
        /// <param name="data"></param>
        /// <param name="signature"></param>
        /// <returns></returns>
        public bool Validate(string timeStamp, string data, string signature)
        {
            var hash = MD5.Create();
            //拼接签名数据
            var signStr = timeStamp + data;
            //将字符串中字符按升序排序
            var sortStr = string.Concat(signStr.OrderBy(c => c));
            var bytes = Encoding.UTF8.GetBytes(sortStr);
            //使用32位大写 MD5签名  
            var md5Val = hash.ComputeHash(bytes);
            var result = new StringBuilder();
            foreach (var c in md5Val) result.Append(c.ToString("X2"));
            var s = result.ToString().ToUpper();
            //与前端传过来的签名参数进行比对
            return  s == signature;
        }
        #endregion
    }
}