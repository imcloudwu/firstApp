using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using Aspose.Cells;

namespace AsposeUse
{
    public partial class Form1 : Form
    {
        String path = @"D:\\testExcel.xls";
        Dictionary<string, string> mapDict = new Dictionary<string, string>();
        
        public Form1()
        {
            InitializeComponent();

        }

        private void Form1_Load(object sender, EventArgs e)
        {
            /*
            for (int i = 0; i < 10; i++)
            {
                row = new DataGridViewRow();
                row.CreateCells(dataGridView1);
                row.Cells[0].Value = "第" + i + "項名稱";
                row.Cells[1].Value = "第" + i + "項內容";
                dataGridView1.Rows.Add(row);
            }
            */
        }

        private void button1_Click(object sender, EventArgs e)
        {
            commitExcel();
        }

        private void setValue()
        {
            mapDict.Clear();
            foreach (DataGridViewRow row in dataGridView1.Rows)
            {
                //???
                if (row.IsNewRow) continue;
                if (row.Cells[0].Value != null && row.Cells[1].Value != null)
                {
                    String key = row.Cells[0].Value.ToString();
                    String value = row.Cells[1].Value.ToString();
                    mapDict.Add(key, value);
                }
            }
        }

        private void commitExcel()
        {
            setValue();
            Aspose.Cells.Workbook workbook = new Aspose.Cells.Workbook();
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;
            int count = 0;

            foreach (KeyValuePair<String, String> s in mapDict)
            {

                for (int i = 0; i <= Convert.ToInt32(s.Value); i++)
                {
                    if (i == 0)
                    {
                        cells[i, count].PutValue(s.Key);
                    }
                    else
                    {
                        cells[i, count].PutValue("[[" + s.Key + i + "]]");
                    }
                }
                count++;
            }

            
            try
            {
                workbook.Save(path);
                label1.Text = "輸出到:";
                linkLabel1.Text = path;
                System.Diagnostics.Process.Start(path);
            }
            catch
            {
                MessageBox.Show("請關閉現有的" + path);
            }
            
            
        }

        private void commitWord()
        {
            setValue();
            Aspose.Words.Document doc = new Aspose.Words.Document(new MemoryStream(Properties.Resources.space));
            Aspose.Words.DocumentBuilder bu = new Aspose.Words.DocumentBuilder(doc);

            bu.StartTable();

            foreach (KeyValuePair<String, String> k in mapDict)
            {
                bu.InsertCell();
                bu.Write(k.Key);
                bu.InsertCell();
                bu.Write(k.Value);
                bu.EndRow();
            }



            foreach (KeyValuePair<String, String> k in mapDict)
            {
                string str = string.Format("MERGEFIELD {0}" + @"\* MERGEFORMAT ", k.Value + "Item " + "value");
                string str1 = string.Format("«{0}»", k.Key + "Item " + "key");
                bu.InsertCell();
                bu.InsertField(str, str1);
            }



            bu.EndRow();

            String path = "D:\\akbbb8.doc";
            doc.Save(path);
            System.Diagnostics.Process.Start(path);


        }

        private void check(object sender, DataGridViewCellValidatingEventArgs e)
        {

            if (dataGridView1.Rows[e.RowIndex].IsNewRow) return; //判斷編輯行[如果是新記錄則退出]

            decimal dci;

            if (e.ColumnIndex == 1)   //指定datagridview 需限制輸入的列
            {

                if (e.FormattedValue != null && e.FormattedValue.ToString().Length > 0)   //确定活動的單元格是否有值輸入
                {

                    if (!decimal.TryParse(e.FormattedValue.ToString(), out dci) || dci < 0)  //判斷輸入的值是否為數值
                    {

                        e.Cancel = true;   //取消驗證事件作業,返回輸入狀態.

                        MessageBox.Show("請輸入數字", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);

                    }

                }

            }
        }

        private void gotoLink(object sender, LinkLabelLinkClickedEventArgs e)
        {
            System.Diagnostics.Process.Start(this.linkLabel1.Text);
        }

        private void button2_Click(object sender, EventArgs e)
        {
            if(File.Exists(path))
            {
                try
                {
                    File.Delete(path);
                    label1.Text = "已刪除" + path;
                    linkLabel1.Text = "";
                }
                catch
                {
                    MessageBox.Show("請關閉現有的" + path);
                }

                
            }
        }


    }
}
