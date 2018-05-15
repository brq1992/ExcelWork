using System;
using System.Windows.Forms;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System.IO;

namespace ExcelTimeCheck
{
    public partial class Form1 : Form
    {
        private string xlsxFileOpenPath;
        private string xlsxFileOutputPath;
        private string sheetName = "打卡时间";
        private int firstColume = 0;
        private int lastColume = 0;
        public Form1()
        {
            InitializeComponent();
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
            if (!string.IsNullOrEmpty(xlsxFileOpenPath) && File.Exists(xlsxFileOpenPath) && !string.IsNullOrEmpty(sheetName))
            {
                FileStream adcFileStream = new FileStream(xlsxFileOpenPath, FileMode.Open);
                XSSFWorkbook sourceWorkbook = new XSSFWorkbook(adcFileStream);
                XSSFSheet sourceSheet = (XSSFSheet)sourceWorkbook.GetSheet(sheetName);
                if (sourceSheet == null)
                {
                    MessageBox.Show("猪，这个表名不对啊！！！快去看看有没有这个表名！！");
                }
            }
        }


        private void sheetNameTB_TextFinished(object sender, EventArgs e)
        {
            
        }

        private void firstCloumeTB_TextChanged(object sender, EventArgs e)
        {
            string input = firstCloumeTB.Text;
            int output;
            bool value = int.TryParse(input, out output);
            if (value)
            {
                firstColume = output;
            }
            else
            {
                MessageBox.Show("输入数字，你是猪么！！");
            }
        }

        private void lastColumeTB_TextChanged(object sender, EventArgs e)
        {
            string input = lastColumeTB.Text;
            int output;
            bool value = int.TryParse(input, out output);
            if (value)
            {
                lastColume = output;
            }
            else
            {
                MessageBox.Show("输入数字，你是猪么！！");
            }
        }
#endregion

        private void button1_Click_1(object sender, EventArgs e)
        {
            FileStream adcFileStream = new FileStream(xlsxFileOpenPath, FileMode.Open);
            XSSFWorkbook sourceWorkbook = new XSSFWorkbook(adcFileStream);
            XSSFSheet sourceSheet = (XSSFSheet)sourceWorkbook.GetSheet(sheetName);
            XSSFWorkbook destWorkbook = new XSSFWorkbook();
            XSSFSheet destWorkSheet = (XSSFSheet)destWorkbook.CreateSheet(sheetName);
            destWorkSheet.CreateRow(sourceSheet.PhysicalNumberOfRows);
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
                            dCell.SetCellValue(sCell.StringCellValue);
                            if (j >= firstColume)
                            {
                                try
                                {
                                    string value = dCell.StringCellValue;
                                    //todo:算法修改，只检测时间00：00
                                    string[] records = System.Text.RegularExpressions.Regex.Split(value, @"\s{2,}");
                                    if (records.Length > 3)
                                    {
                                        string result = string.Format("{0}  \n{1}", records[0], records[records.Length - 2]);
                                        dCell.SetCellValue(result);
                                        ICellStyle styeCellStyle = destWorkbook.CreateCellStyle();
                                        styeCellStyle.FillForegroundColor = 57;
                                        styeCellStyle.FillPattern = FillPattern.SolidForeground;
                                        dCell.CellStyle = styeCellStyle;
                                    }
                                }
                                catch (Exception ex)
                                {
                                    Console.WriteLine("Exception: " + ex.Message);
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
        }
        
    }
}
