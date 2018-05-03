using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.OleDb;
using System.IO;
using Microsoft.Office.Interop.Excel;
using System.Xml.Linq;

namespace ExcelForXml
{
    public partial class Form1 : Form
    {
        private List<string> m_pPathList = new List<string>();
        private Dictionary<string, Dictionary<string, List<string>>> m_pPathKeyDicMul = new Dictionary<string, Dictionary<string, List<string>>>();
        string path = System.IO.Directory.GetCurrentDirectory();
        int m_iProgressValue = 10;
        public Form1()
        {
            InitializeComponent();
            CreatePathList();
        }
        private void listBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            m_pPathList.Clear();
            listBox1.SelectionMode = SelectionMode.MultiExtended;
            foreach (var item in listBox1.SelectedItems)
            {
                m_pPathList.Add(item.ToString());
            }

        }
        private void OpenExcel(string str_key)
        {
            string strFileName = path + "/excel/" + str_key + ".xlsx";
            object missing = System.Reflection.Missing.Value;
            Dictionary<string, List<string>> m_pPathKeyDic = new Dictionary<string, List<string>>();
            Microsoft.Office.Interop.Excel.Application excel = new Microsoft.Office.Interop.Excel.Application();//lauch excel application
            if (excel == null)
            {
                MessageBox.Show("cant find excel :"+ strFileName);
            }
            else
            {
                excel.Visible = false; excel.UserControl = true;
                // 以只读的形式打开EXCEL文件
                Workbook wb = excel.Application.Workbooks.Open(strFileName, missing, true, missing, missing, missing,
                 missing, missing, missing, true, missing, missing, missing, missing, missing);
                //取得第一个工作薄
                Worksheet ws = (Worksheet)wb.Worksheets.get_Item(1);
                //取得总记录行数   (包括标题列)
                int rowsint = ws.UsedRange.Cells.Rows.Count; //得到行数
                int colsint = ws.UsedRange.Cells.Columns.Count; //得到行数
                for (int i = 1; i <= rowsint; i++)
                {
                    for (int n = 1; n <= colsint; n++)
                    {
                        m_iProgressValue++;
                        string key = ((Range)ws.UsedRange.Cells[1, n]).Text;
                        if (i == 1)
                            m_pPathKeyDic.Add(key, new List<string>());
                        else
                            m_pPathKeyDic[key].Add(((Range)ws.UsedRange.Cells[i, n]).Text);
                    }
                }
            }
            excel.Quit(); excel = null;
            GC.Collect();
            if (m_pPathKeyDicMul.ContainsKey(str_key))
                m_pPathKeyDicMul[str_key] = m_pPathKeyDic;
            else
                m_pPathKeyDicMul.Add(str_key, m_pPathKeyDic);
        }

        public void CreatXml(string table_name)
        {
            Dictionary<string, List<string>> m_pPathKeyDic = m_pPathKeyDicMul[table_name];
            XDocument xDoc = new XDocument();
            XElement xRot = new XElement(table_name);
            xDoc.Add(xRot);
            string ele_name = table_name.Replace("Table", "");
            int iCount = 0;
            foreach (var key in m_pPathKeyDic.Keys)
            {
                iCount = m_pPathKeyDic[key].Count;
                break;
            }
            for (int i = 0; i < iCount; i++)
            {
                XElement xEle = new XElement(ele_name);
                foreach (var key in m_pPathKeyDic.Keys)
                {
                    m_iProgressValue++;
                    XAttribute xAtt = new XAttribute(key, m_pPathKeyDic[key][i]);//ID_item,name
                    xEle.Add(xAtt);
                }
                xRot.Add(xEle);
            }
            xDoc.Save(path+"/xml/"+ table_name+".xml");
        }

        private void label1_Click(object sender, EventArgs e)
        {
            
        }

        private void CreatePathList()
        {
            listBox1.Items.Clear();
            DirectoryInfo di = new DirectoryInfo(path);
            FileInfo[] fis = di.GetFiles("*.xlsx", SearchOption.AllDirectories);
            for (int i = 0; i < fis.Length; i++)
            {
                string aFirstName = fis[i].Name.Substring(fis[i].Name.LastIndexOf("\\") + 1, (fis[i].Name.LastIndexOf(".") - fis[i].Name.LastIndexOf("\\") - 1));
                listBox1.Items.Add(aFirstName);
            }
        }
        private void button1_Click(object sender, EventArgs e)
        {
            m_iProgressValue = 50;
            progressBar1.Minimum = m_iProgressValue;
            progressBar1.Maximum = 300;
            m_pPathKeyDicMul.Clear();
            SetProgressValue(m_iProgressValue);
            for (int i = 0; i < m_pPathList.Count; i++)
            {
                m_iProgressValue += 1;
                m_pPathKeyDicMul.Add(m_pPathList[i],new Dictionary<string, List<string>>());
                OpenExcel(m_pPathList[i]);
                CreatXml(m_pPathList[i]);
                SetProgressValue(m_iProgressValue);
            }
            //CreatXml();
            SetProgressValue(progressBar1.Maximum);
            MessageBox.Show("转换完成");

        }

        void SetProgressValue(int value)
        {
            if (value > progressBar1.Maximum)
                progressBar1.Value = progressBar1.Maximum;
            else
                progressBar1.Value = value;
        }
        private void progressBar1_Click(object sender, EventArgs e)
        {

        }
    }
}
