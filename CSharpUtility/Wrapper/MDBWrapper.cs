using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data.OleDb;
using System.Collections;
using System.Data;

namespace CSharpUtility.Wrapper
{
    // 使用例
    //ArrayList alSearch = new ArrayList();
    //using (MDBController mdb = new MDBController())
    //{
    //    mdb.Open(sDBPath);
    //    alSearch = mdb.Execute(query);
    //}

    /// -----------------------------------------------------------------------------
    /// <summary>
    /// MDBを操作するラッパー
    /// </summary>
    /// -----------------------------------------------------------------------------
    class MDBWrapper
    {
        //DBオブジェクト
        private OleDbConnection DBConnection;
        private OleDbCommand DBCommand;
        private string DBConnectionString = "";

        /// -----------------------------------------------------------------------------
        /// <summary>
        ///     コンストラクタ
        /// </summary>
        /// -----------------------------------------------------------------------------
        public MDBWrapper()
        {
            this.DBConnection = new OleDbConnection();
            this.DBConnectionString += "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=";
        }

        /// -----------------------------------------------------------------------------
        /// <summary>
        ///     MDBを開く
        /// </summary>
        /// <param name="path">DBパス</param>
        /// -----------------------------------------------------------------------------
        public void Open(string path)
        {
            this.DBConnection.ConnectionString = this.DBConnectionString + @path;
            this.DBConnection.Open();
        }

        /// -----------------------------------------------------------------------------
        /// <summary>
        ///     MDBを閉じる
        /// </summary>
        /// -----------------------------------------------------------------------------
        public void Close()
        {
            this.DBConnection.Close();
        }

        /// -----------------------------------------------------------------------------
        /// <summary>
        /// クエリを実行する
        /// </summary>
        /// <param name="query">クエリ</param>
        /// <returns>実行結果</returns>
        /// -----------------------------------------------------------------------------
        public ArrayList Execute(string query)
        {
            DataTable dtTbl = new DataTable();
            this.DBCommand = new OleDbCommand(query, this.DBConnection);

            OleDbDataAdapter dbAdp = new OleDbDataAdapter(this.DBCommand);
            dbAdp.Fill(dtTbl);

            ArrayList alResult = new ArrayList();
            string[] sColuumn = new string[dtTbl.Columns.Count];

            for (int i = 0; i < dtTbl.Rows.Count; i++)
            {
                for (int j = 0; j < dtTbl.Columns.Count; j++)
                {
                    sColuumn[j] = dtTbl.Rows[i][j].ToString();
                }

                alResult.Add(sColuumn);
            }

            return alResult;
        }

        /// -----------------------------------------------------------------------------
        /// <summary>
        /// 終了
        /// </summary>
        /// -----------------------------------------------------------------------------
        public void Dispose()
        {
            //終了・破棄
            this.DBCommand.Dispose();
            this.DBConnection.Close();
            this.DBConnection.Dispose();
        }
    }
}
