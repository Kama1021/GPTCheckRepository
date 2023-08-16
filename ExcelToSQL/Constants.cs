namespace ExcelToSQL
{
    public static class Constants
    {
        public const int ACQUISITION_SHEET_NUM = 1;   //取得するシートの番号

        public const int COLUMN_NAME_ROW = 1;   //カラム名を取得する行
        public const int COLUMN_FORMAT_ROW = 2;   //カラムのフォーマットを取得する行

        public const int SLEEP_INTERVAL_COUNT = 1;   //非同期処理でループ中にスリープする時間
        public const int SLEEP_BEFORE_RETURN_COUNT = 1000;   //非同期処理でreturn前にスリープする時間

        public const string CELL_DELIMITER = "!";   //セル同士の区切り文字
        public const string ROW_CARRIAGE_RETURN_CAHR = "?";   //ロウの区切り文字

        public const string ERROR_TEXT = "error";   //エラー時に関数が返す値
    }
}
