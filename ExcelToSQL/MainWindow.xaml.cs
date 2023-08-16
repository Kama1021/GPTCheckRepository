using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using ClosedXML.Excel;
using DocumentFormat.OpenXml.Spreadsheet;
using System.IO;
using DocumentFormat.OpenXml.Drawing;
using Microsoft.Win32;

namespace ExcelToSQL
{
    public partial class MainWindow : Window
    {
        private int loadedCell = 0;   //読み込み済みのセルの数を格納する数値
        private int cellLimit =0;   //セルの最大数を格納する数値
        private int lastRow = 0;   //読み込んだシートの記述がある最後の行を格納する数値
        private int lastColumn = 0;   //読み込んだシートの記述がある最後の列を格納する数値
        private bool isLoadingFile = false;   //ファイルを読み込み中か
        private bool isCanLoad = false;   //ファイルを読み込めたか
       

        private string[]? dragFilePaths =null;   //ドラッグされてきたファイルパスを格納する
        private string? dropFilePaths =null;   //ドロップされたファイルパスを格納する
        private string? fileNameWithoutExtension =null;   //テーブル名を格納する


        private List <string> columnList = new List<string>();   //取得したエクセルファイルの1行目(カラム名)を取得する用のリスト
        private List <string> columnFormatList = new List<string>();   //取得したエクセルファイルの2行目からカラムのフォーマットを取得する用のリスト

        //クラスインスタンス作成
        private GuessDataTypeSystem guessDataTypeSystem = new GuessDataTypeSystem();

        public MainWindow()
        {
            InitializeComponent();
            textGrid.IsHitTestVisible = true;   //グリッド内の接触系イベントを可能状態にする
        }


        /// <summary>
        /// ファイルがウィンドウに侵入したときの処理
        /// </summary>
        private void WindowDragEnter(object sender, DragEventArgs e)
        {
            textGrid.IsHitTestVisible = false;   //接触系イベント不可にする(こうしないとテキストボックスにドロップできない)
        }

        /// <summary>
        /// ファイルがウィンドウから出た時の処理
        /// </summary>
        private void WindowDragLeave(object sender, DragEventArgs e)
        {
            textGrid.IsHitTestVisible = true;   //接触系イベント可能にする
        }

        /// <summary>
        /// ファイルがウィンドウ内でドラッグされている間の処理
        /// </summary>
        private void WindowDragOver(object sender, DragEventArgs e)
        {
            //ドラッグされたオブジェクトがファイルかどうかチェック
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                //ファイルパスを取得(複数受け取った場合も考えて配列)
                dragFilePaths = e.Data.GetData(DataFormats.FileDrop) as string[];
                //ファイルパスがnullでなく、かつ1つ以上の場合
                if (dragFilePaths != null && dragFilePaths.Length > 0)
                {
                    //取得したファイルの最初の要素の拡張子を小文字で取得
                    var dragFileExtension = System.IO.Path.GetExtension(dragFilePaths[0])?.ToLower();
                    //ドラッグされたオブジェクトがエクセルファイルであれば
                    if (dragFileExtension == ".xlsx")
                    {
                        e.Effects = DragDropEffects.Copy;   //ドラッグの制御をコピーモードにする
                        fileNameWithoutExtension = System.IO.Path.GetFileNameWithoutExtension(dragFilePaths[0]);
                        isCanLoad = true;
                    }
                    //そうでなければ
                    else
                    {
                        e.Effects = DragDropEffects.None;   //ドラッグの制御をドロップ不可にする
                        isCanLoad = false;
                        
                    }
                }
            }
            //ファイル以外がドロップされた場合
            else
            {
                e.Effects = DragDropEffects.None;   //ドラッグの制御をドロップ不可にする
                isCanLoad = false;
            }
        }

        

        /// <summary>
        /// ファイルがウィンドウ内でドロップ(手放された)時の処理
        /// </summary>
        private async void WindowDrop(object sender, DragEventArgs e)
        {
            //エラーで処理したい状態の場合
            if (dragFilePaths == null || isCanLoad == false)
            {
                ConsoleText.Text = "取得不可能なファイル形式でした\n";
                return;
            }
            //読み込み中の場合
            if (isLoadingFile == true)
            {
                ConsoleText.Text = "現在のファイルが読み込み完了するまで新しいファイルの読み込みはできません\n";
                return;
            }
            //ファイル名(テーブル名)がバリデーションチェックに引っかかったら
            if (!Validation.IsValidPostgreSQLIdentifier(fileNameWithoutExtension))
            {
                ConsoleText.Text = "テーブル名が不正です。\nテーブル名はファイル名から取得されます。";
                return;
            }
            //初期化処理
            ConsoleText.Text = string.Empty;
            cellLoadProgresPercentText.Text="0%";
            ConsoleText.Text = "";

            isLoadingFile = true;   //読み込み中フラグを立てる
            
            dropFilePaths = dragFilePaths[0];   //ドラッグ時に取得したファイルパスのうち、先頭のものだけドロップ用のファイルパスに格納

            // タイマーを設定して読み込み中のインジケータを表示
            int dotCount = 0;
            progressGrid.Visibility = Visibility.Visible;
            codeBox.Text= string.Empty;
            
            //DispacherTimer(一定間隔事にイベントを発生させるタイマー)を0.5秒間隔に設定して作成
            System.Windows.Threading.DispatcherTimer timer = new System.Windows.Threading.DispatcherTimer{Interval = TimeSpan.FromSeconds(0.5)};
            //DispatcherTimerのイベントハンドラを作成
            timer.Tick += (s, args) =>
            {
                dotCount = (dotCount + 1) % 4;   //カウンタを1ずつ加算し、4になったら0に戻す
                codeBox.Text = "読み込み中" + new string('.', dotCount);   //"."をカウント分描画する
                cellLoadProgresText.Text = loadedCell+ " / " +cellLimit;   //読み込み済みのセル/すべてのセル表示のテキストを描画する
                cellLoadProgressBar.Value = loadedCell;   //プログレスバーの進度を描画する
                //除算時にエラーが出ないように0チェックを行う
                if (cellLimit!=0&&loadedCell!=0)
                {
                    cellLoadProgresPercentText.Text = (((float)loadedCell / cellLimit) *100).ToString("F1") + "% ";   //読み込み状況を少数第二位切り捨ての％のテキストで表示する
                }

            };
            timer.Start();   //タイマーの稼働を開始

            //非同期処理開始
            string[,]? result = await Task.Run(() => LoadExcelData(dropFilePaths));
            
            timer.Stop();   //非同期処理が完了した時点でタイマーをストップ

            textGrid.IsHitTestVisible = true;   //接触系イベント可能にする
            progressGrid.Visibility = Visibility.Hidden;   //プログレスバー周りを非表示にする
            isLoadingFile = false;   //読み込み中フラグを折る
            if (result == null)
            {

                ConsoleText.Text = "不正なカラム名です。";
                return;   //例外処理
            }
            if (result[0,0] == Constants.ERROR_TEXT)
            {
                // エラーメッセージを表示
                ConsoleText.Text = "致命的なエラーが発生しました。\n他のプロセスでファイルを開いている可能性が高いです";
                return;
            }
            string sqlCode = ConvertExcelToSQL(result);   //Excelから取得した文字列をSQL文に変換する
            codeBox.Text = sqlCode;

        }


        /// <summary>
        /// Excelのデータを読み込むための関数
        /// </summary>
        /// <param name="filePaths">読み込むべきエクセルファイルのパス</param>
        /// <returns>ストリングビルダーでエクセルファイルを展開したもの</returns>
        private string[,]? LoadExcelData(string filePaths)
        {
            try
            {
                //Excelファイルを適切にクローズするためにusingを使用
                using (XLWorkbook excelFile = new XLWorkbook(@filePaths)) // ファイルパスからエクセルファイルを取得
                {
                    IXLWorksheet workSheet = excelFile.Worksheet(Constants.ACQUISITION_SHEET_NUM); // エクセルファイルの指定されたシートを取得
                    lastRow = workSheet.LastRowUsed().RowNumber(); // 記述がある最後の行を取得
                    lastColumn = workSheet.LastColumnUsed().ColumnNumber(); // 記述がある最後の列を取得
                    string[,] cellTable = new string[lastRow, lastColumn];
                    cellLimit = lastRow * lastColumn; // セルの総数を取得
                    loadedCell = 0; // 読み込み済みのセルを初期化
                    columnList = new List<string>();
                    columnFormatList = new List<string>();

                    // await中にコントロールの操作を行う部分(この中じゃないとエラーが発生する)
                    Application.Current.Dispatcher.Invoke(() =>
                    {
                        cellLoadProgressBar.Maximum = cellLimit; // プログレスバーの最大値設定
                    });

                    // 行分処理
                    for (int i = 1; i <= lastRow; i++)
                    {
                        // 列分処理
                        for (int j = 1; j <= lastColumn; j++)
                        {
                            IXLCell cell = workSheet.Cell(i, j); // ()内の座標のセルを取得
                            cellTable[i - 1, j - 1] = cell.GetFormattedString();
                            
                            string cellValue = cell.GetFormattedString();

                            // 時刻かどうかチェック
                            TimeSpan timeValue;
                            if (TimeSpan.TryParse(cellValue, out timeValue))
                            {
                                cellValue = timeValue.ToString(@"hh\:mm\:ss");   //時刻の場合12時間表記を修正する
                            }
                            loadedCell++; // 読み込み済みのセルを加算
                                          // 1行目の時にカラム名を取得
                            if (i == Constants.COLUMN_NAME_ROW)
                            {
                                // カラム名に対するバリデーションチェック
                                if (Validation.IsValidPostgreSQLIdentifier(cell.GetFormattedString()) == false)
                                {
                                    return null;
                                }
                                columnList.Add(cell.GetFormattedString());
                            }
                            // 2列目の時にカラムのフォーマットを取得
                            else if (i == Constants.COLUMN_FORMAT_ROW)
                            {
                                columnFormatList.Add(guessDataTypeSystem.GuessDataType(cell.GetFormattedString()));
                            }
                        }
                    }

                    System.Threading.Thread.Sleep(Constants.SLEEP_BEFORE_RETURN_COUNT); // プログレスバーを100％にするための間隔
                    return cellTable;
                }
            }
            catch (IOException e)
            {
                // 必要に応じてnullを返すなどの処理
                return new string[1, 1] { { Constants.ERROR_TEXT } };
            }
        }


        /// <summary>
        /// 渡されたExcelデータからSQL文を作成する
        /// </summary>
        /// <param name="excelData"></param>
        /// <returns></returns>
        private string ConvertExcelToSQL(string[,] excelData)
        {
            StringBuilder sqlCode= new StringBuilder();   //ストリングビルダーのインスタンスを作成
            //テーブル作成部
            sqlCode.AppendLine("CREATE TABLE " + fileNameWithoutExtension + "(") ;   //テーブル作成宣言
            for (int i = 0; i < columnList.Count; i++)
            {
                sqlCode.AppendLine(columnList[i] +" " + columnFormatList[i]+",");   //カラムの宣言
            }
            sqlCode.Length -= 3;   //末尾を1文字削除
            sqlCode.AppendLine();   //改行
            sqlCode.AppendLine(");");   //テーブル作成閉じ

            //要素作成部
            for (int i = 1; i < lastRow-1; i++)
            {
                sqlCode.Append($"INSERT INTO {fileNameWithoutExtension} (");
                for (int j = 0; j < lastColumn; j++)
                {
                    sqlCode.Append($"{columnList[j]},");
                }
                sqlCode.Length -= 1;
                sqlCode.Append(") \nVALUES(");
                for(int k = 0;k < lastColumn; k++)
                {
                    sqlCode.Append($"'{excelData[i,k]}',");
                }
                sqlCode.Length -= 1;
                sqlCode.AppendLine(");");
            }



            return sqlCode.ToString();
        }

        private void ClickCopyButton(object sender, RoutedEventArgs e)
        {
            Clipboard.SetText(codeBox.Text);
            ConsoleText.Text = "コピーできeました！";
        }

        private void ClickExportButton(object sender, RoutedEventArgs e)
        {
            var saveFileDialog = new SaveFileDialog();

            saveFileDialog.Filter = "SQLファイル|*.sql";
            saveFileDialog.FileName = fileNameWithoutExtension;

            if (saveFileDialog.ShowDialog()==true)
            {
                File.WriteAllText(saveFileDialog.FileName, codeBox.Text);
            }
            ConsoleText.Text = "ファイルの保存が成功しました。\n"+saveFileDialog.FileName;
        }

    }
}
