using System;
using System.Text.RegularExpressions;

namespace ExcelToSQL
{
    public static class Validation
    {

        public static bool IsValidPostgreSQLIdentifier(string name)
        {
            // 長さの制限をチェック
            if (name.Length > 63)
            {
                return false;
            }

            // 数字で始まるかチェック
            if (char.IsDigit(name[0]))
            {
                return false;
            }

            // 英数字とアンダースコア以外の文字が含まれていないかチェック
            if (!Regex.IsMatch(name, @"^[a-zA-Z_][a-zA-Z0-9_]*$"))
            {
                return false;
            }

            // 予約語のチェック（必要に応じて追加・変更）
            string[] reservedWords = { "select", "table", "from", /* 他の予約語 */ };
            if (Array.Exists(reservedWords, word => word.Equals(name, StringComparison.OrdinalIgnoreCase)))
            {
                return false;
            }

            return true;
        }

    }
}
