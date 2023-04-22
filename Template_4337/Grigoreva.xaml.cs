using Microsoft.Office.Interop.Excel;
using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;
using Excel = Microsoft.Office.Interop.Excel;

namespace Template_4337
{
    /// <summary>
    /// Логика взаимодействия для Grigoreva.xaml
    /// </summary>
    public partial class Grigoreva : System.Windows.Window
    {
        public Grigoreva()
        {
            InitializeComponent();
        }

        private void btn1_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog()
            {
                DefaultExt = "*.xls;*.xlsx",
                Filter = "файл Excel (Spisok.xlsx)|*.xlsx",
                Title = "Выберите файл базы данных"
            };
            if (!(ofd.ShowDialog() == true))
                return;
            string[,] list;
            Excel.Application ObjWorkExcel = new Excel.Application();
            Excel.Workbook ObjWorkBook = ObjWorkExcel.Workbooks.Open(ofd.FileName);
            Excel.Worksheet ObjWorkSheet = (Excel.Worksheet)ObjWorkBook.Sheets[1];
            var lastCell = ObjWorkSheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell);
            int _columns = (int)lastCell.Column;
            int _rows = (int)lastCell.Row;
            list = new string[_rows, _columns];
            for (int j = 0; j < _columns; j++)
                for (int i = 0; i < _rows; i++)
                    list[i, j] = ObjWorkSheet.Cells[i + 1, j + 1].Text;
            using(PROBEntities1 db = new PROBEntities1())
            {
                for (int i = 1; i < 71; i++)
                {
                    /*db.Clients.Add(new Clients()
                    {
                        FIO = list[i, 0],
                        Code = list[i, 1],
                        Birthday = DateTime.ParseExact(list[i, 2].ToString(), "dd.mm.yyyy", System.Globalization.CultureInfo.InvariantCulture),
                        Index = list[i, 3],
                        City = list[i, 4],
                        Street = list[i, 5],
                        House = int.Parse(list[i, 6]),
                        Flat = int.Parse(list[i, 7]),
                        Email = list[i, 8]
                    });*/
                    db.Street.Add(new Street()
                    {
                        Street1 = list[i, 5]
                    });
                }
                db.SaveChanges();
                MessageBox.Show("Данные добавлены");
                ObjWorkBook.Close(false, Type.Missing, Type.Missing);
                ObjWorkExcel.Quit();
                GC.Collect();
            }
        }

        private void btn2_Click(object sender, RoutedEventArgs e)
        {
            List<Clients> allClients;
            List<string> allStreets;
            using (PROBEntities1 db = new PROBEntities1())
            {
                allClients = db.Clients.ToList().OrderBy(s => s.FIO).ToList();
                var query = from s in db.Street select s.Street1;
                allStreets = query.Distinct().ToList();
            }
            var app = new Excel.Application();
            app.SheetsInNewWorkbook = allStreets.Count();
            Excel.Workbook workbook = app.Workbooks.Add(Type.Missing);
            for (int i = 0; i < allStreets.Count(); i++)
            {
                int startRowIndex = 1;
                Excel.Worksheet worksheet = app.Worksheets.Item[i + 1];
                worksheet.Name = allStreets[i];
                worksheet.Cells[1][startRowIndex] = "Порядковый номер";
                worksheet.Cells[2][startRowIndex] = "ФИО";
                startRowIndex++;
                var clientsCategories = allClients.GroupBy(s => s.Street).ToList();
                foreach (var c in clientsCategories)
                {
                    if (c.Key == allStreets[i])
                    {
                        Excel.Range headerRange = worksheet.Range[worksheet.Cells[1][1], worksheet.Cells[2][1]];
                        headerRange.Merge();
                        headerRange.Value = allStreets[i];
                        headerRange.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                        headerRange.Font.Italic = true;
                        startRowIndex++;
                        foreach (Clients c1 in allClients)
                        {
                            if (c1.Street == c.Key)
                            {
                                worksheet.Cells[1][startRowIndex] = c1.Code;
                                worksheet.Cells[2][startRowIndex] = c1.FIO;
                                startRowIndex++;
                            }
                        }
                        worksheet.Cells[1][startRowIndex].Formula = $"=СЧЁТ(A3:A{startRowIndex - 1})";
                        worksheet.Cells[1][startRowIndex].Font.Bold = true;
                    }
                    else
                    {
                        continue;
                    }
                }
                Excel.Range rangeBorders = worksheet.Range[worksheet.Cells[1][1], worksheet.Cells[2][startRowIndex - 1]];
                rangeBorders.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = rangeBorders.Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = rangeBorders.Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle =
                    rangeBorders.Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = rangeBorders.Borders[Excel.XlBordersIndex.xlInsideHorizontal].LineStyle =
                    rangeBorders.Borders[Excel.XlBordersIndex.xlInsideVertical].LineStyle = Excel.XlLineStyle.xlContinuous;
                worksheet.Columns.AutoFit();
            }
            app.Visible = true;
        }
    }
}
