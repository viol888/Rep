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
using Excel = Microsoft.Office.Interop.Excel;
using System.Data.SqlClient;
using System.Configuration;
using System.IO;
using System.Net;
using System.Net.Http;

namespace Report
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
        private static readonly HttpClient client = new HttpClient();
        private void button1_Click(object sender, EventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog();
            ofd.DefaultExt = "*.xls;*.xlsx";
            ofd.Filter = "Microsoft Excel (*.xls*)|*.xls*";
            ofd.Multiselect = true;
            if (ofd.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                foreach (string fileName in ofd.FileNames)
                {
                    textBox1.Text = textBox1.Text + fileName + "\n";
                }
            }
            else
                MessageBox.Show("Вы не выбрали файл для открытия", "Загрузка данных...", MessageBoxButtons.OK, MessageBoxIcon.Error);            
            Excel.Application xlApp = new Excel.Application();            
            string[] Individal_Runs = textBox1.Text.Split('\n').Where(s => !string.IsNullOrEmpty(s)).ToArray();
            foreach (string s in Individal_Runs)
            {             
                FillExcel(xlApp,s);
            }
                xlApp.Quit();
                System.Runtime.InteropServices.Marshal.ReleaseComObject(xlApp);
                GC.Collect();            
            }
        public string SendRequest(string s)
        {
            WebRequest request = WebRequest.Create("https://npchk.nalog.ru/chk-lst.html");
            // Set the Method property of the request to POST.
            request.Method = "POST";
            // Create POST data and convert it to a byte array.
            string lst = "lst="+s;
            byte[] byteArray = Encoding.UTF8.GetBytes(lst);
            // Set the ContentType property of the WebRequest.
            request.ContentType = "application/x-www-form-urlencoded";
            // Set the ContentLength property of the WebRequest.
            request.ContentLength = byteArray.Length;
            // Get the request stream.
            Stream dataStream = request.GetRequestStream();
            // Write the data to the request stream.
            dataStream.Write(byteArray, 0, byteArray.Length);
            // Close the Stream object.
            dataStream.Close();
            // Get the response.
            WebResponse response = request.GetResponse();
            // Display the status.
            Console.WriteLine(((HttpWebResponse)response).StatusDescription);
            // Get the stream containing content returned by the server.
            dataStream = response.GetResponseStream();
            // Open the stream using a StreamReader for easy access.
            StreamReader reader = new StreamReader(dataStream);
            // Read the content.
            string responseFromServer = reader.ReadToEnd();
            // Display the content.
            //MessageBox.Show(responseFromServer);
            // Clean up the streams.
            reader.Close();
            dataStream.Close();
            response.Close();
            return (responseFromServer);
        }

        public void FillExcel( Excel.Application xlApp, string s)
        {
            Excel.Workbook xlWorkBook = xlApp.Workbooks.Open(s);
            Excel.Worksheet xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(2);
            object misValue = System.Reflection.Missing.Value;
            var xlSheets = xlWorkBook.Sheets as Excel.Sheets;
            var xlNewSheet = (Excel.Worksheet)xlSheets.Add(misValue, xlWorkBook.Worksheets[xlWorkBook.Worksheets.Count],
            misValue, misValue);
            string stringForSend = "";
            string[] massWithAnswer;
            int massIndex = 0;
            int LastRow = xlWorkSheet.UsedRange.Rows[xlWorkSheet.UsedRange.Rows.Count].Row - 9;
            xlWorkSheet.Range["C14:C" + LastRow].Copy(); //копируем диапазон ячеек
            xlNewSheet.Range["A14:A" + LastRow].PasteSpecial();
            xlWorkSheet.Range["I14:I" + LastRow].Copy(); //копируем диапазон ячеек
            xlNewSheet.Range["B14:B" + LastRow].PasteSpecial();
            xlWorkSheet.Range["J14:J" + LastRow].Copy(); //копируем диапазон ячеек
            xlNewSheet.Range["C14:C" + LastRow].PasteSpecial();
            xlNewSheet.Columns["A:A"].ColumnWidth = 15;
            xlNewSheet.Columns["B:B"].ColumnWidth = 25;
            xlNewSheet.Columns["C:C"].ColumnWidth = 15;
            xlNewSheet.Columns["F:F"].ColumnWidth = 16;
            xlNewSheet.Columns["G:G"].ColumnWidth = 16;
            xlNewSheet.Columns["H:H"].ColumnWidth = 12;
            xlWorkSheet.Cells[9, 3].Copy(); //копируем диапазон ячеек
            xlNewSheet.Cells[9, 1].PasteSpecial();
            xlNewSheet.Cells[2, 10].Value = "Признаки, проставляемые при проверке контрагентов";
            xlNewSheet.Cells[2, 10].Font.Bold = true;
            xlNewSheet.Cells[3, 10].Value = "0 - Налогоплательщик зарегистрирован в ЕГРН и имел статус действующего в указанную дату";
            xlNewSheet.Cells[3, 10].Font.Bold = false;
            xlNewSheet.Cells[4, 10].Value = "1 - Налогоплательщик зарегистрирован в ЕГРН, но не имел статус действующего в указанную дату";
            xlNewSheet.Cells[5, 10].Value = "2 - Налогоплательщик зарегистрирован в ЕГРН";
            xlNewSheet.Cells[6, 10].Value = "3 - Налогоплательщик с указанным ИНН зарегистрирован в ЕГРН, КПП не соответствует ИНН или не указан";
            xlNewSheet.Cells[7, 10].Value = "4 - Налогоплательщик с указанным ИНН не зарегистрирован в ЕГРН";
            xlNewSheet.Cells[8, 10].Value = "5 - Некорректный ИНН";
            xlNewSheet.Cells[9, 10].Value = "6 - Недопустимое количество символов ИНН";
            xlNewSheet.Cells[10, 10].Value = "7 - Недопустимое количество символов КПП";
            xlNewSheet.Cells[11, 10].Value = "8 - Недопустимые символы в ИНН";
            xlNewSheet.Cells[12, 10].Value = "9 - Недопустимые символы в КПП";
            xlNewSheet.Cells[13, 10].Value = "10 - КПП не должен использоваться при проверке ИП";
            xlNewSheet.Cells[14, 10].Value = "11 - некорректный формат даты";
            xlNewSheet.Cells[15, 10].Value = "12 - некорректная дата(ранее 01.01.1991 или позднее текущей даты)";
            xlNewSheet.Cells[16, 10].Value = "? -ошибка обработки запроса";
            xlNewSheet.Cells[6, 1].Font.Size = 14;
            xlNewSheet.Cells[6, 1].Font.Name = "Times New Roman";
            xlNewSheet.Cells[6, 1].Value = "КНИГА ПРОДАЖ";
            xlNewSheet.Range["A6:C6"].Merge();
            xlNewSheet.Range["A6:C6"].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            xlNewSheet.Cells[6, 1].Font.Bold = true;
            xlNewSheet.Cells[9, 1].Font.Bold = true;
            xlNewSheet.Cells[9, 1].Font.Underline = true;
            xlNewSheet.Cells[9, 1].Font.Italic = true;
            int row = 31;
            while (row != LastRow)
            {
                xlNewSheet.Cells[row, 4].Formula = "= CLEAN(C" + row + ")";
                xlNewSheet.Cells[row, 5].Formula = "= FIND(\"/\",C" + row + ",1)";
                xlNewSheet.Cells[row, 6].NumberFormat = "#";
                xlNewSheet.Cells[row, 6].Formula = "= IF(ISERROR(E" + row + "),IF(D" + row + "=\"\",\"\",IF(MID(C" + row + ",1,1)= \"0\",D" + row + ", D" + row + " * 1)), " +
                    "IF(MID(C" + row + ",1,1)= \"0\", MID(D" + row + ",1,E" + row + "-1),MID(D" + row + ",1,E" + row + "-1)*1))";
                xlNewSheet.Cells[row, 7].Formula = "= IF(ISERROR(E" + row + "),\"\",IF(MID(C" + row + ",1,1)= \"0\",MID(D" + row + ",E" + row + "+ 1,9),MID(D" + row + ",E" + row + " +1,9)*1))";
                xlNewSheet.Cells[row, 8].Formula = "=MID(A" + row + ",FIND(\"от \", A" + row + ", 1)+3, 10)";
                //Funk("6658122658	665801001	09.01.2019\n6658122658  665812265  09.01.2019");
                stringForSend = stringForSend + xlNewSheet.Cells[row, 6].Value + "  " + xlNewSheet.Cells[row, 7].Value + "  " + xlNewSheet.Cells[row, 8].Value + "\n";
                xlNewSheet.Cells[row, 9].Value = "0";
                row++;
                //xlNewSheet.Cells[i, 9].Value = (Funk(xlNewSheet.Cells[i, 6].Value + "  " + xlNewSheet.Cells[i, 7].Value + "  " + xlNewSheet.Cells[i, 8].Value)).Split(new char[] { ' ' }, StringSplitOptions.RemoveEmptyEntries)[0];
                if (row % 100 == 0)
                {
                    massWithAnswer = SendRequest(stringForSend).Split('\n');
                    for (massIndex=0; massIndex < massWithAnswer.Length - 1; massIndex++)
                    {
                        if (massWithAnswer[massIndex].Split(' ')[0] != "0")
                            xlNewSheet.Cells[row - 10 + massWithAnswer.Length - 1, 9].Value = massWithAnswer[massIndex].Split(' ')[0];
                    }
                    stringForSend = "";
                }                
            }
            //= НАЙТИ("/"; C32; 1)
            // xlWorkBook.Save();
            massWithAnswer = SendRequest(stringForSend).Split('\n');
            int t= massWithAnswer.Length - 1;
            for (int n = 0; n < massWithAnswer.Length - 1; n++)
            {
                if (massWithAnswer[n].Split(' ')[0] != "0")
                    xlNewSheet.Cells[row - t, 9].Value = massWithAnswer[n].Split(' ')[0];
                massIndex++;
                t--;
            }
            stringForSend = "";
            massWithAnswer = null;
            xlWorkBook.Close(true, misValue, misValue);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(xlWorkBook);
            xlWorkBook = null;
            MessageBox.Show("файл " + s + " преобразован");
        }
    }
}
