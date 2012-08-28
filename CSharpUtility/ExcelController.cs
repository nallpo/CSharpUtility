using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Runtime.InteropServices;
using System.Collections;

namespace CSharpUtility
{
    //  要:参照の追加
    //  
    //  COM:Microsoft Excel 9.0 Object

    //  *使用例*
    //
    //  ExcelController ec = new ExcelController("hello.xls");  //宣言、ファイルを指定
    //  ec.setWorkSheet("sheet1");                              //ワークシートを設定
    //
    //  int r = ec.searchRows("hello", 1);                      //文字列を検索し、該当する行を取得
    //  Console.WriteLine( ec.getCell(r, 2) );                  //指定したセルの値を取得
    //
    //  ec.Close();                                             //終了処理

    class ExcelController
    {
        //xlsファイルパス
        private string exPath;

        //Excelオブジェクト
        private Excel.Application oXls;


        //workbookオブジェクト
        Excel.Workbook oWBook;

        //Worksheetオブジェクト
        Excel.Worksheet oSheet;

        /// <summary>
        /// エクセルデータの操作を行う
        /// </summary>
        /// <param name="path">xlsファイルパス</param>
        public ExcelController(string path)
        {
            exPath = path;

            oXls = new Excel.Application();

            //Excel画面を表示しない
            oXls.Visible = false;

            //Excelファイルをオープンする
            oWBook = (Excel.Workbook)(oXls.Workbooks.Open(exPath));
        }

        /// <summary>
        /// 終了処理を行う
        /// </summary>
        public void Close()
        {
            Marshal.ReleaseComObject(oSheet);
            oWBook.Close(false, false, Type.Missing);
            oXls.Quit();
            Marshal.ReleaseComObject(oWBook);
            Marshal.ReleaseComObject(oXls);
        }

        /// <summary>
        /// ワークシートを設定する。
        /// </summary>
        /// <param name="wsName">ワークシート名</param>
        public void setWorkSheet(string wsName)
        {
            oSheet = (Excel.Worksheet)oWBook.Sheets[getSheetIndex(wsName, oWBook.Sheets)];
        }

        //指定されたワークシート名のインデックスを返す
        private int getSheetIndex(string sheetName, Excel.Sheets shs)
        {
            int i = 0;
            foreach (Excel.Worksheet sh in shs)
            {
                if (sheetName == sh.Name)
                {
                    return i + 1;
                }
                i += 1;
            }
            return 0;
        }

        /// <summary>
        /// ワークシートから指定した文字列を検索し、該当する行番号を返す。
        /// 該当しない場合、-1を返す。
        /// </summary>
        /// <param name="str">文字列</param>
        /// <param name="columnNum">検索列</param>
        /// <returns>該当する行番号</returns>
        public int searchRows(string str, int columnNum)
        {
            string column = getColumnName(columnNum);

            //セルの範囲を指定
            Excel.Range range = oSheet.get_Range(column + "1:" + column + oSheet.UsedRange.Rows.Count);

            //結果収納配列
            ArrayList al = new ArrayList();
            int re = -1;

            //検索
            foreach (Excel.Range r in range)
            {
                string c = r.Text.ToString();

                if (System.Text.RegularExpressions.Regex.IsMatch(c, "^" + str + "$"))
                {
                    re = r.Row;
                    break;
                }
            }

            return re;
        }

        /// <summary>
        /// ワークシートから指定したセルの文字列を取得する。
        /// </summary>
        /// <param name="row">行</param>
        /// <param name="column">列</param>
        /// <returns>文字列</returns>
        public string getCell(int row, int column)
        {
            Excel.Range rng;

            rng = (Excel.Range)oSheet.Cells[row, column];

            return rng.Text.ToString();
        }

        /// <summary>
        /// 列番号からアルファベットを取得する
        /// </summary>
        /// <param name="index">列</param>
        /// <returns>アルファベット</returns>
        public string getColumnName(int index)
        {
            index--;
            string str = "";
            do
            {
                str = Convert.ToChar(index % 26 + 0x41) + str;
            } while ((index = index / 26 - 1) != -1);

            return str;
        }
    }
}
