using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using BPM;
using BPM.Server;
using NPOI;
using NPOI.HSSF;
using NPOI.HSSF.UserModel;
using System.Data;
using System.Data.SqlClient;

namespace GetProcesserFromExcel
{
    class Program
    {
        static void Main(string[] args)
        {
        }
        public static void getProcesser()
        {
            Context context = Context.Current;
            FlowDataTable table = context.FormDataSet.Tables["TestFlow"];//表单中的表名称
            string fileid = Convert.ToString(table.Rows[0]["Attachment"]);//获取附件ID；
            string file = System.AppDomain.CurrentDomain.BaseDirectory + @"Attachments\" + Attachment.FileIDToPath(fileid);//获取服务器端系统中存附件的路径
            string files = System.AppDomain.CurrentDomain.BaseDirectory + @"sAttachments\" + fileid + ".xls";//获取新建路径sAttachments（前提是已经有创建好此目录）
            System.IO.File.Copy(file, files, true);//复制到新的路径；

            List<String> name = new List<String>();//定义处理人姓名列表集合

            FileStream fs = File.OpenRead(files);//打开文件
            HSSFWorkbook book = new HSSFWorkbook(fs);
            HSSFSheet sheet = (HSSFSheet)book.GetSheetAt(0);
            for (int r = 1; r <= sheet.LastRowNum; r++)
            {
                HSSFRow row = (HSSFRow)sheet.GetRow(r);
                HSSFCell cellname = (HSSFCell)row.GetCell(1);//认为处理人在excel中位于B列，可根据实际情况修改
                if (!name.Contains(Convert.ToString(cellname.StringCellValue)))
                {
                    name.Add(Convert.ToString(cellname.StringCellValue));
                }

            }

            //把集合中的处理人姓名拼接成“A,B,C”格式
            string displayname = "";
            foreach (string disname in name)
            {
                displayname += "'" + disname + "',";
            }
            displayname = displayname.Substring(0, displayname.Length - 1);

            string sqlStr = string.Format(@"select Account  from BPMSysUsers where DisplayName in ({0})", displayname);
            string fzaccount = "";
            using (IDataReader reader = context.IDataProvider.ExecuteReader(sqlStr.ToString()))
            {
                while (reader.Read())
                {
                    fzaccount += Convert.ToString(reader["Account"]) + ",";
                }
            }
            fzaccount = fzaccount.Substring(0, fzaccount.Length - 1); //把集合中的处理人账号拼接成“A,B,C”格式
            //以上都是从excel中读取数据，然后获取账号

            //以下是根据账号返回流程处理人
            if (fzaccount.Contains(","))//说明有多个处理人
            {
                MemberCollection members = new MemberCollection();
                string[] value = fzaccount.Split(',');
                for (int i = 0; i < value.Length; i++)
                {
                    string account = Convert.ToString(value[i]);
                    if (!members.ContainsUser(account))
                        members.Add(Member.FromAccount(account));
                }
                return members;
            }
            else
            {
                return Member.FromAccount(fzaccount);
            }
        }
    }
}
