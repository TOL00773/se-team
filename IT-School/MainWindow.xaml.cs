using System;
using System.Collections;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using Excel = Microsoft.Office.Interop.Excel;

namespace IT_School
{
    /// <summary>
    /// Логика взаимодействия для MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {

        public class Organization
        {
            public string FirstName { get; set; }
            public string LastName { get; set; }
            public string FirstName { get; set; }
            public string LastName { get; set; }
            public string FirstName { get; set; }
            public string LastName { get; set; }
            public string FirstName { get; set; }
            public string LastName { get; set; }
        }

        ObservableCollection<Organization> custdata = new ObservableCollection<Organization>();

        public MainWindow()
        {
            InitializeComponent();

            //GetData() creates a collection of Customer data from a database


            Organization a = new Organization();
            a.FirstName = "Sasha";
            a.LastName = "Sasha";
            custdata.Add(a);
            //Bind the DataGrid to the customer data
            DG1.DataContext = custdata;
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {

        }

        private void Button_Click_1(object sender, RoutedEventArgs e)
        {
            string filename = "data.xls";
            string adress = "http://www.admoblkaluga.ru/opendata/4027064263-ObrazovanieReestrUchrejd/data-124-structure-1.xls";
            WebClient myWebClient = new WebClient();
            myWebClient.DownloadFile(adress, filename);
        }
        private static string filename = Directory.GetCurrentDirectory()+@"\data.xls";
        private void Button_Click_2(object sender, RoutedEventArgs e)
        {

            Excel.Application ObjExcel = new Excel.Application();
            //Открываем книгу.                                                                                                                                                        
            Excel.Workbook ObjWorkBook = ObjExcel.Workbooks.Open(filename, 0, true, 5, "", "", false, Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);
            //Выбираем таблицу(лист).
            Excel.Worksheet ObjWorkSheet;
            ObjWorkSheet = (Excel.Worksheet)ObjWorkBook.Sheets[1];
            string info = "";
            //Выбираем первые сто записей из столбца.
            //Выбираем область таблицы. (в нашем случае просто ячейку)
            for (int i = 1; i < 501; i++)
            {
                //Выбираем область таблицы. (в нашем случае просто ячейку)
                Microsoft.Office.Interop.Excel.Range range = ObjWorkSheet.get_Range("B" + i.ToString(), "D" + i.ToString());
                //Добавляем полученный из ячейки текст.
                info = range.Text.ToString();
                Organization a = new Organization();
                a.FirstName = info;
                custdata.Add(a);

            }
        }
    }
}
