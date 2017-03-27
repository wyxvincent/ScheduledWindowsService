using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace Scheduled_Task
{
    class CSVHelper
    {
        /// <summary>
        /// 导出报表为Csv(客户端用 WinForm、WPF或Windows服务等)
        /// </summary>
        /// <param name="dt">DataTable</param>
        /// <param name="strFilePath">物理路径</param>
        /// <param name="tableheader">表头</param>
        /// <param name="columname">字段标题,逗号分隔</param>
        public static bool dt2csvClient(DataTable dt, string strFilePath)
        {
            try
            {
                StringBuilder sb = new StringBuilder();

                foreach (DataColumn dc in dt.Columns)
                {
                    sb.Append(dc.ColumnName.TrimStart().TrimEnd());
                    sb.Append(",");
                }
                sb.Remove(sb.ToString().LastIndexOf(","), 1);
                sb.Append("\r\n");

                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    for (int j = 0; j < dt.Columns.Count; j++)
                    {
                        sb.Append("\"" + dt.Rows[i][j].ToString().TrimStart().TrimEnd() + "\"");
                        sb.Append(",");
                    }
                    sb.Remove(sb.ToString().LastIndexOf(","), 1);
                    sb.Append("\r\n");
                }
                using (System.IO.StreamWriter sw = new System.IO.StreamWriter(strFilePath, false, System.Text.Encoding.UTF8))
                {
                    sw.Write(sb.ToString());
                }
                return true;
            }
            catch
            {
                return false;
            }

        }

        /// <summary>
        /// 将Csv读入DataTable
        /// </summary>
        /// <param name="filePath">csv文件路径</param>
        /// <param name="n">表示第n行是字段title,第n+1行是记录开始</param>
        public static DataTable csv2dt(string filePath, int n, DataTable dt)
        {
            StreamReader reader = new StreamReader(filePath, System.Text.Encoding.UTF8, false);
            int i = 0, m = 0;
            reader.Peek();
            while (reader.Peek() > 0)
            {
                m = m + 1;
                string str = reader.ReadLine();
                if (m >= n + 1)
                {
                    string[] split = str.Split(',');

                    System.Data.DataRow dr = dt.NewRow();
                    for (i = 0; i < split.Length; i++)
                    {
                        dr[i] = split[i];
                    }
                    dt.Rows.Add(dr);
                }
            }
            return dt;
        }

        /// <summary>
        /// 将CSV读入DataTable(版本2)
        /// </summary>
        /// <param name="fileContent"></param>
        /// <returns></returns>
        public static DataTable CSVToDataTable(byte[] fileContent)
        {
            string s = string.Empty;
            string src = string.Empty;
            SortedList sl = new SortedList();

            String[] split = null;
            DataTable table = new DataTable("auto");
            DataRow row = null;

            Stream stream = new MemoryStream(fileContent);
            StreamReader fs = new StreamReader(stream, Encoding.GetEncoding("gb2312"));

            //创建与数据源对应的数据列 
            s = fs.ReadLine();
            split = s.Split(',');
            foreach (String colname in split)
            {
                StringBuilder sbColName = new StringBuilder();
                string[] colnames = colname.Split(' ');
                for (int i = 0; i < colnames.Length; i++)
                {
                    if (string.IsNullOrEmpty(colnames[i]))
                        continue;
                    sbColName.Append(colnames[i]);
                }
                table.Columns.Add(sbColName.ToString(), System.Type.GetType("System.String"));
            }
            //将数据填入数据表 
            int j = 0;
            s = string.Empty;
            while (!string.IsNullOrEmpty(s += fs.ReadLine()))
            {
                //if (!s.EndsWith(","))
                //{
                //    continue;
                //}
                if (!string.IsNullOrEmpty(s))
                {
                    j = 0;
                    row = table.NewRow();

                    src = s.Replace("\"\"", "'");
                    MatchCollection col = Regex.Matches(src, ",\"([^\"]+)\"", RegexOptions.ExplicitCapture);
                    IEnumerator ie = col.GetEnumerator();
                    while (ie.MoveNext())
                    {
                        string patn = ie.Current.ToString();
                        int key = src.Substring(0, src.IndexOf(patn)).Split(',').Length;
                        if (!sl.ContainsKey(key))
                        {
                            sl.Add(key, patn.Trim(new char[] { ',', '"' }).Replace("'", "\""));
                            src = src.Replace(patn, ",");
                        }
                    }

                    string[] arr = src.Split(',');
                    for (int i = 0; i < arr.Length; i++)
                    {
                        if (!sl.ContainsKey(i))
                            sl.Add(i, arr[i]);
                    }
                    //foreach (String colname in arr)
                    //{
                    //    row[j] = colname.Replace("@:::::", ",");
                    //    j++;
                    //}

                    IDictionaryEnumerator ienum = sl.GetEnumerator();
                    while (ienum.MoveNext())
                    {
                        row[j] = ienum.Value.ToString().Replace("'", "\"");
                        j++;
                    }
                    table.Rows.Add(row);
                    sl.Clear();
                    s = string.Empty;
                    src = string.Empty;
                }
            }
            fs.Close();
            return table;
        }
    }
}
