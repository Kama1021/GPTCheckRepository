using System;
using System.Globalization;
using System.Net;

namespace ExcelToSQL
{
    class GuessDataTypeSystem
    {
        /// <summary>
        /// 渡された文字列から、代替のSQLフォーマットを判別する関数
        /// </summary>
        /// <param name="data">判別したい文字列</param>
        /// <returns>フォーマット名をstringで返す</returns>
        public string GuessDataType(string data)
        {
            // 整数型
            if (short.TryParse(data, out _))
                return "smallint";
            if (int.TryParse(data, out _))
                return "integer";
            if (long.TryParse(data, out _))
                return "bigint";

            // 不動少数点数型
            if (float.TryParse(data, out _))
                return "real";
            if (double.TryParse(data, out _))
                return "double precision";
            if (decimal.TryParse(data, out _))
                return "numeric";

            // ブール型
            if (bool.TryParse(data, out _))
                return "boolean";

            // 時間型
            TimeSpan timeValue;
            if (TimeSpan.TryParse(data, out timeValue))
            {
                data = timeValue.ToString(@"hh\:mm\:ss"); // 24時間制に変換
                return "time";
            }
            // 日付と時刻の形式 "2023-06-24 00:00:00"
            if (DateTime.TryParseExact(data, "yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture, DateTimeStyles.None, out _))
                return "timestamp with time"; 

            // 日付の形式 "2023/6/27"
            if (DateTime.TryParseExact(data, "yyyy/M/d", CultureInfo.InvariantCulture, DateTimeStyles.None, out _))
                return "timestamp without time";

            if (DateTime.TryParse(data, out _))
                return "timestamp";


            // アドレス型
            if (IPAddress.TryParse(data, out _))
                return "inet";

            // UUID
            if (Guid.TryParse(data, out _))
                return "uuid";

            //DEFAULT
            return "text";
        }
    }
}
