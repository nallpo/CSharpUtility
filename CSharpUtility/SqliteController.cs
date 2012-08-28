using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data.SQLite;

namespace CSharpUtility
{
    //  *使用例 読み込み*
    //
    //  SqliteController sc = new SqliteController("hello.db");
    //
    //  sc.open();
    //  sc.setSql("select * from test order by hoge desc limit 3");
    //
    //  SQLiteDataReader reader = sc.read();
    //
    //  while (reader.Read())
    //  {
    //      Console.WriteLine("hoge:" + reader[0] + ", mage:" + reader[1]);
    //  }
    //  sc.close();

    //  *使用例 書き込み*
    //
    //  SqliteController sc = new SqliteController("hello.db");
    //
    //  sc.open();
    //  sc.setSql("insert into test (hoge, mage) values('foo', 'bar');");
    //  sc.execute();
    //  sc.close();

    class SqliteController
    {
        private SQLiteConnection con;
        private SQLiteCommand cmd;
        SQLiteDataReader reader;

        public SqliteController(string path)
        {
            con = new SQLiteConnection("Data Source=" + path);
            cmd = con.CreateCommand();
        }

        public void open()
        {
            con.Open();
        }

        public void close()
        {
            con.Close();
        }

        public SQLiteDataReader read()
        {
            reader = cmd.ExecuteReader();
            return reader;
        }

        public void execute()
        {
            cmd.ExecuteNonQuery();
        }

        public void setSql(string sql)
        {
            if (reader != null)
            {
                reader.Close();
            }

            cmd.CommandText = sql;
        }
    }
}
