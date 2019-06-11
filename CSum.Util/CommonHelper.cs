using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Text;

namespace CSum.Util
{
    /// <summary>
    ///     常用公共类
    /// </summary>
    public class CommonHelper
    {
        #region 删除数组中的重复项

        /// <summary>
        ///     删除数组中的重复项
        /// </summary>
        /// <param name="values"></param>
        /// <returns></returns>
        public static string[] RemoveDup(string[] values)
        {
            var list = new List<string>();
            for (var i = 0; i < values.Length; i++) //遍历数组成员
            {
                if (!list.Contains(values[i])) list.Add(values[i]);
                ;
            }

            return list.ToArray();
        }

        #endregion

        #region 自动生成日期编号

        /// <summary>
        ///     自动生成编号  201008251145409865
        /// </summary>
        /// <returns></returns>
        public static string CreateNo()
        {
            var random = new Random();
            var strRandom = random.Next(1000, 10000).ToString(); //生成编号 
            var code = DateTime.Now.ToString("yyyyMMddHHmmss") + strRandom; //形如
            return code;
        }

        #endregion

        #region 生成0-9随机数

        /// <summary>
        ///     生成0-9随机数
        /// </summary>
        /// <param name="codeNum">生成长度</param>
        /// <returns></returns>
        public static string RndNum(int codeNum)
        {
            var sb = new StringBuilder(codeNum);
            var rand = new Random();
            for (var i = 1; i < codeNum + 1; i++)
            {
                var t = rand.Next(9);
                sb.AppendFormat("{0}", t);
            }

            return sb.ToString();
        }

        #endregion

        #region 生成指定范围内的不重复随机数

        /// <summary>
        ///     生成指定范围内的不重复随机数
        /// </summary>
        /// <param name="Number">随机数量</param>
        /// <param name="minNum">下限值</param>
        /// <param name="maxNum">上限值</param>
        /// <returns></returns>
        public int[] GetRandomArray(int Number, int minNum, int maxNum)
        {
            int j;
            var b = new int[Number];
            var r = new Random();
            for (j = 0; j < Number; j++)
            {
                var i = r.Next(minNum, maxNum + 1);
                var num = 0;
                for (var k = 0; k < j; k++)
                    if (b[k] == i)
                        num = num + 1;
                if (num == 0)
                    b[j] = i;
                else
                    j = j - 1;
            }

            return b;
        }

        #endregion

        #region Stopwatch计时器

        /// <summary>
        ///     计时器开始
        /// </summary>
        /// <returns></returns>
        public static Stopwatch TimerStart()
        {
            var watch = new Stopwatch();
            watch.Reset();
            watch.Start();
            return watch;
        }

        /// <summary>
        ///     计时器结束
        /// </summary>
        /// <param name="watch"></param>
        /// <returns></returns>
        public static string TimerEnd(Stopwatch watch)
        {
            watch.Stop();
            double costtime = watch.ElapsedMilliseconds;
            return costtime.ToString();
        }

        #endregion

        #region 删除最后一个字符之后的字符

        /// <summary>
        ///     删除最后结尾的一个逗号
        /// </summary>
        public static string DelLastComma(string str)
        {
            return str.Substring(0, str.LastIndexOf(","));
        }

        /// <summary>
        ///     删除最后结尾的指定字符后的字符
        /// </summary>
        public static string DelLastChar(string str, string strchar)
        {
            return str.Substring(0, str.LastIndexOf(strchar));
        }

        /// <summary>
        ///     删除最后结尾的长度
        /// </summary>
        /// <param name="str"></param>
        /// <param name="Length"></param>
        /// <returns></returns>
        public static string DelLastLength(string str, int Length)
        {
            if (string.IsNullOrEmpty(str))
                return "";
            str = str.Substring(0, str.Length - Length);
            return str;
        }

        #endregion
    }
}