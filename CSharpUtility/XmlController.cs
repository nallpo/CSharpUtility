using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.Xml;
using System.Collections;

namespace CSharpUtility
{
    //  *使用例 読み込み*
    //
    //  XmlController xc = new XmlController("hello.xml");                              //宣言、ファイルを指定
    //
    //  xc.openReader();                                                                //読み取り開始
    //
    //  string[] node = { "setting", "input" };                                         //ノードを階層順に指定
    //  string[][] atribute = new string[][] { new string[] { "code", "1234567" } };    //属性値を指定
    //  Console.WriteLine( xc.getElementValue(node, atribute) );                        //該当する値を取得
    //
    //  xc.closeReader();                                                               //読み取り終了


    //  *使用例 書き込み*
    //
    //  XmlController xc = new XmlController("hello.xml");                              //宣言、ファイルを指定
    //
    //  XmlData xd = new XmlData("setting");                                            //XMLデータ格納用インスタンスの宣言(引数はノード名)
    //
    //  XmlData xd_input = xd.setChild("input");                                        //子ノードの設定
    //  xd_input.setsetAttributes("code", "1234567");                                   //属性値の設定
    //  xd_input.setValue("hello");                                                     //値の設定
    //
    //  xc.openWriter();                                                                //書き込み開始
    //  xc.Save(xd);                                                                    //書き込み処理
    //  xc.closeWriter();                                                               //書き込み終了
    
    class XmlController
    {
        //xmlファイルパス
        private string xmlPath;

        //ファイルストリーム
        private FileStream fs;

        //Reader, Writer
        private XmlTextReader reader;
        private XmlTextWriter writer;

        /// <summary>
        /// XMLデータの操作を行う
        /// </summary>
        /// <param name="path">xmlファイルパス</param>
        public XmlController(string path)
        {
            xmlPath = path;

            //空ファイルの作成
            if (!File.Exists(xmlPath))
            {
                StreamWriter toucher = new StreamWriter(xmlPath);
                toucher.Close();
            }
        }

        /// <summary>
        /// XmlTextReaderを開く
        /// </summary>
        public void openReader()
        {
            fs = new FileStream(xmlPath, FileMode.Open, FileAccess.Read);
            reader = new XmlTextReader(fs);
        }

        /// <summary>
        /// XmlTextWriterを開く
        /// </summary>
        public void openWriter()
        {
            fs = new FileStream(xmlPath, FileMode.Create, FileAccess.Write);
            writer = new XmlTextWriter(fs, System.Text.Encoding.Default);
        }

        /// <summary>
        /// XmlTextReaderを終了させる
        /// </summary>
        public void closeReader()
        {
            reader.Close();
            fs.Close();
        }

        /// <summary>
        /// XmlTextWriterを終了させる
        /// </summary>
        public void closeWriter()
        {
            writer.Close();
            fs.Close();
        }

        /// <summary>
        /// xmlから指定したノード名、属性名・値に該当するノードの値を返す
        /// 該当しない場合、nullを返す。
        /// </summary>
        /// <param name="elements">ノード名(階層順に配列で指定)</param>
        /// <param name="attributes">属性({{属性名, 属性値}, {属性名, 属性値}...}となる二次元配列)</param>
        /// <returns>ノードの値</returns>
        public string getElementValue(string[] elements, string[][] attributes = null)
        {
            //ノード関係
            ArrayList nodes = new ArrayList();

            //ノードの深さ
            int nodeDepth = 0;

            //検索
            try
            {
                while (reader.Read())
                {
                    reader.MoveToContent();

                    //ノード要素にデータがある場合
                    if (reader.NodeType == XmlNodeType.Element)
                    {
                        // 開始タグを発見した場合
                        if (reader.NodeType == XmlNodeType.Element)
                        {
                            //ノードが浅くなった場合
                            if (nodeDepth >= reader.Depth && nodes.Count > 0)
                            {
                                //同じノード数になるまで調整
                                for (int i = nodeDepth; i >= reader.Depth; i--)
                                {
                                    nodes.RemoveAt(nodes.Count - 1);
                                }
                            }

                            //ノードの深さ
                            nodeDepth = reader.Depth;

                            //ノード名
                            nodes.Add(reader.LocalName);

                            //ノードの確認
                            string[] tNodes = new string[nodes.Count];
                            for (int i = 0; i < nodes.Count; i++)
                            {
                                tNodes[i] = nodes[i].ToString();
                            }

                            //検索対象のノードである場合
                            if (tNodes.SequenceEqual(elements))
                            {
                                //属性を確認しない場合
                                if (attributes == null)
                                {
                                    if (reader.ReadString() == null)
                                    {
                                        return "";
                                    }

                                    return reader.ReadString();
                                }

                                //属性を確認する場合
                                if (reader.HasAttributes)
                                {
                                    int count = 0;

                                    // すべての属性を表示
                                    for (int i = 0; i < reader.AttributeCount; i++)
                                    {
                                        // 属性ノードへ移動
                                        reader.MoveToAttribute(i);

                                        for (int j = 0; j < attributes.Length; j++)
                                        {
                                            if (new string[] { reader.Name, reader.Value }.SequenceEqual(attributes[j]))
                                            {
                                                count++;
                                                break;
                                            }
                                        }
                                    }

                                    if (count == reader.AttributeCount)
                                    {
                                        if (reader.ReadString() == null)
                                        {
                                            return "";
                                        }

                                        return reader.ReadString();
                                    }

                                    // すべての属性を出力したら、元のノード(エレメントノード)に戻る
                                    reader.MoveToElement();
                                }
                            }
                        }
                    }
                }
            }
            catch (XmlException e)
            {
                Console.WriteLine(e);
            }

            reader.ResetState();
            return "";
        }

        /// <summary>
        /// XmlDataで作成されたXMLデータからXMLファイルを作成する。
        /// </summary>
        /// <param name="data">XMLデータデータ</param>
        public void Save(XmlData data)
        {
            //XMLファイルにインデントを入れる
            writer.Formatting = Formatting.Indented;

            writer.WriteStartDocument(true);

            //データの格納
            data.writeXml(writer);

            writer.WriteEndDocument();
        }
    }

    public class XmlData
    {
        //ノード名
        private string nodeName;

        //ノード値
        private string nodeValue;

        //子ノード
        private Dictionary<string, XmlData> nodeChild = new Dictionary<string, XmlData>();

        //属性値
        private Hashtable nodeAttributes = new Hashtable();

        /// <param name="name">ノード名</param>
        public XmlData(string name)
        {
            nodeName = name;
        }

        /// <summary>
        /// XmlDataで作成されたXMLデータからXMLファイルを作成する。
        /// </summary>
        /// <param name="w">XmlTextWriterインスタンス</param>
        public void writeXml(XmlTextWriter w)
        {
            //ノードを作成
            w.WriteStartElement(nodeName);

            if(nodeChild.Count > 0)
            {
                foreach (KeyValuePair<string, XmlData> kv in nodeChild)
                {
                    kv.Value.writeXml(w);
                }
            }

            if (nodeAttributes != null)
            {
                foreach (string key in nodeAttributes.Keys)
                {
                    //ノード属性を出力
                    w.WriteAttributeString(key, (string)nodeAttributes[key]);
                }
            }

            w.WriteString(nodeValue);

            //ノードを閉じる
            w.WriteEndElement();
        }

        /// <summary>
        /// ノード名を取得する
        /// </summary>
        /// <returns>ノード名</returns>
        public string getName()
        {
            return nodeName;
        }

        /// <summary>
        /// 指定した子ノードを取得する
        /// </summary>
        /// <param name="name">ノード名</param>
        /// <returns>子ノード</returns>
        public XmlData getChild(string name)
        {
            return (XmlData)nodeChild[name];
        }

        /// <summary>
        /// 指定したノードの値を取得する
        /// </summary>
        /// <param name="name">ノード名</param>
        /// <returns>ノード値</returns>
        public string getValue(string name)
        {
            return nodeValue;
        }

        /// <summary>
        /// ノードの属性値を取得する
        /// </summary>
        /// <param name="name">属性名</param>
        /// <returns>属性値</returns>
        public string getAttributes(string name)
        {
            return (string)nodeAttributes[name];
        }

        /// <summary>
        /// ノード名を設定する
        /// </summary>
        /// <param name="name">ノード名</param>
        public void setName(string name)
        {
            nodeName = name;
        }

        /// <summary>
        /// 子ノードを設定する
        /// </summary>
        /// <param name="name">ノード名</param>
        /// <returns>子ノード</returns>
        public XmlData setChild(string name)
        {
            nodeChild.Add(name, new XmlData(name));
            return (XmlData)nodeChild[name];
        }

        /// <summary>
        /// ノード値を設定する
        /// </summary>
        /// <param name="name">ノード値</param>
        public void setValue(string value)
        {
            nodeValue = value;
        }

        /// <summary>
        /// ノードの属性値を設定する
        /// </summary>
        /// <param name="name">属性名</param>
        /// <param name="value">属性値</param>
        public void setAttributes(string name, string value)
        {
            nodeAttributes.Add(name, value);
        }
    }
}
