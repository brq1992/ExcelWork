using System;
using System.Collections.Generic;
using System.Windows.Forms;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;

namespace ExcelTimeCheck
{
    public partial class Form1 : Form
    {
        private string xlsxFileOpenPath;
        private string xlsxFileOutputPath;
        private string sheetName = "打卡时间";
        private int firstColume = 0;
        private int lastColume = 0;

        string[] char26 = { "a", "b", "c", "d", "e", "f", "g", "h", "i", "j", "k", "l", "m", "n", "o", "p", 
                         "q", "r", "s", "t", "u", "v", "w", "x", "y", "z"};

        public Form1()
        {
            InitializeComponent();
            sheetNameTB.Text = sheetName;
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            
        }



#region label
        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void label2_Click(object sender, EventArgs e)
        {

        }

        private void label5_Click(object sender, EventArgs e)
        {

        }
#endregion


#region btn

        private void button1_Click(object sender, EventArgs e)
        {
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                openPathLabel.Text = openFileDialog1.FileName;
                xlsxFileOpenPath = openFileDialog1.FileName;
                if (!File.Exists(xlsxFileOpenPath))
                {
                    MessageBox.Show("找不到这个文件，猪！！重新指定！！");
                    openPathLabel.Text = "你是猪么！";
                    xlsxFileOpenPath = string.Empty;
                }
            }
        }

        private void saveBtn_Click(object sender, EventArgs e)
        {
            if (saveFileDialog1.ShowDialog() == DialogResult.OK)
            {
                outputPath.Text = saveFileDialog1.FileName;
                xlsxFileOutputPath = saveFileDialog1.FileName;
            }
        }

#endregion
       

#region textBox
        private void sheetNameTB_TextChanged(object sender, EventArgs e)
        {
            sheetName = sheetNameTB.Text;
            //if (!string.IsNullOrEmpty(xlsxFileOpenPath) && File.Exists(xlsxFileOpenPath) && !string.IsNullOrEmpty(sheetName))
            //{
            //    FileStream adcFileStream = new FileStream(xlsxFileOpenPath, FileMode.Open);
            //    XSSFWorkbook sourceWorkbook = new XSSFWorkbook(adcFileStream);
            //    XSSFSheet sourceSheet = (XSSFSheet)sourceWorkbook.GetSheet(sheetName);
            //    if (sourceSheet == null)
            //    {
            //        MessageBox.Show("猪，这个表名不对啊！！！快去看看有没有这个表名！！");
            //    }
            //}
        }

        private void firstCloumeTB_TextChanged(object sender, EventArgs e)
        {
            //string input = firstCloumeTB.Text;
            //int output;
            //bool value = int.TryParse(input, out output);
            //if (value)
            //{
            //    firstColume = output;
            //}
            //else
            //{
            //    MessageBox.Show("输入数字，你是猪么！！");
            //}
        }

        private void lastColumeTB_TextChanged(object sender, EventArgs e)
        {
            //string input = lastColumeTB.Text;
            //int output;
            //bool value = int.TryParse(input, out output);
            //if (value)
            //{
            //    lastColume = output;
            //}
            //else
            //{
            //    MessageBox.Show("输入数字，你是猪么！！");
            //}
        }
#endregion

        private int GetInter(string str)
        {
            //string input = str;
            //int output;
            //bool value = int.TryParse(input, out output);
            //if (!value)
            //{
            //    MessageBox.Show("输入数字，你是猪么！！");
            //    output = -1;
            //}
            //return output;
            return ToIndex(str);
        }

        public int ToIndex(string columnName)
        {
            if (!Regex.IsMatch(columnName.ToUpper(), @"[A-Z]+"))
            {
                MessageBox.Show("输入的列数不对呀!!猪。要输入 ABCD");
                return -1;
            }
            char[] chars = columnName.ToUpper().ToCharArray();
            int index = chars.Select((t, i) => (t - 'A' + 1) * (int)Math.Pow(26, chars.Length - i - 1)).Sum();
            return index - 1;
        }

        public string ToName(int index)
        {
            if (index < 0) { throw new Exception("invalid parameter"); }
            List<string> chars = new List<string>();
            do
            {
                if (chars.Count > 0) index--;
                chars.Insert(0, ((char)(index % 26 + 'A')).ToString());
                index = (index - index % 26) / 26;
            } 
            while (index > 0);
            return string.Join(string.Empty, chars.ToArray());
        }

        private void button1_Click_1(object sender, EventArgs e)
        {
            button1.Text = "执行中..";
            button1.Enabled = false;
            bool check = true;
            if (!File.Exists(xlsxFileOpenPath))
            {
                MessageBox.Show("找不到这个文件哦~小猪，看看文件路径对不对~");
                button1.Text = "执行";
                button1.Enabled = true;
                return;
            }
            FileStream adcFileStream = new FileStream(xlsxFileOpenPath, FileMode.Open);
            XSSFWorkbook sourceWorkbook = new XSSFWorkbook(adcFileStream);
            XSSFSheet sourceSheet = (XSSFSheet)sourceWorkbook.GetSheet(sheetName);
            if (sourceSheet == null)
            {
                MessageBox.Show("找不到这个表名哦！小猪，检查一下表名对不对：" + sheetName);
                adcFileStream.Close();
                sourceWorkbook.Close();
                button1.Text = "执行";
                button1.Enabled = true;
                return;
            }
            firstColume =  GetInter(firstCloumeTB.Text);
            lastColume = GetInter(lastColumeTB.Text) + 1;
            XSSFWorkbook destWorkbook = new XSSFWorkbook();
            XSSFSheet destWorkSheet = (XSSFSheet)destWorkbook.CreateSheet(sheetName);
            for (int i = 0; i < sourceSheet.PhysicalNumberOfRows; i++)
            {
                IRow row = sourceSheet.GetRow(i);
                IRow dRow = destWorkSheet.CreateRow(i);
                if (row != null)
                {
                    for (int j = 0; j < lastColume; j++)
                    {
                        ICell sCell = row.GetCell(j);
                        ICell dCell = dRow.CreateCell(j);
                        if (sCell != null && dCell != null)
                        {
                            string cellStr = "";
                            try
                            {
                                cellStr = sCell.StringCellValue;
                            }
                            catch (InvalidOperationException exception)
                            {
                                cellStr = "该格异常！";
                            }
                            dCell.SetCellValue(cellStr);
                            if (j >= firstColume)
                            {
                                try
                                {
                                    string value = dCell.StringCellValue;
                                    MatchCollection mc = System.Text.RegularExpressions.Regex.Matches(value, @"\d{2}:\d{2}");
                                    if (mc.Count > 2)
                                    {
                                        string result = string.Format("{0}  \n{1}", mc[0], mc[mc.Count - 1]);
                                        dCell.SetCellValue(result);
                                        ICellStyle styeCellStyle = destWorkbook.CreateCellStyle();
                                        styeCellStyle.FillForegroundColor = 57;
                                        styeCellStyle.FillPattern = FillPattern.SolidForeground;
                                        dCell.CellStyle = styeCellStyle;
                                    }
                                }
                                catch (Exception ex)
                                {
                                    check = false;
                                    MessageBox.Show("执行异常：" + ex.Message);
                                }
                            }
                        }
                    }
                }
            }
            FileStream sw = File.Create(xlsxFileOutputPath);
            destWorkbook.Write(sw);
            sw.Close();
            adcFileStream.Close();
            sourceWorkbook.Close();
            destWorkbook.Close();

            if (check)
            {
                DialogResult dr = MessageBox.Show("是否打开表格", "完成", MessageBoxButtons.OKCancel, MessageBoxIcon.Question);
                if (dr == DialogResult.OK)
                {
                    if (File.Exists(xlsxFileOutputPath))
                    {
                        System.Diagnostics.Process.Start(xlsxFileOutputPath);
                    }
                }
            }
            else
            {
                MessageBox.Show("出错了！！");
            }

            button1.Text = "执行";
            button1.Enabled = true;
        }

        
    }
}
