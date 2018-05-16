using System;
using System.Windows.Forms;

namespace ExcelTimeCheck
{
    static class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new Form1());
            //string[] value = new[]
            //{"09:26  19:33  19:55  ", "09:26  19:33  19:55  19:58", "09:26  19:33  19:55  19:58  你"};
            //for (int i = 0; i < value.Length; i++)
            //{
            //    MatchCollection mc = System.Text.RegularExpressions.Regex.Matches(value[i], @"\d{2}:\d{2}");
            //    Console.WriteLine("re: " + mc.Count);
            //    string[] records = System.Text.RegularExpressions.Regex.Split(value[i], @"\d{2}:\d{2}");
            //}
        }
    }
}
