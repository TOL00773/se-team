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

<<<<<<< HEAD
=======

>>>>>>> 7aba9fad65ebb305604ac0ae921a105771184c52
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
            
            Microsoft.Office.Interop.Excel.Application ObjExcel = new Microsoft.Office.Interop.Excel.Application();
            //Открываем книгу.                                                                                                                                                        
            Microsoft.Office.Interop.Excel.Workbook ObjWorkBook = ObjExcel.Workbooks.Open(filename, 0, true, 5, "", "", false, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);
            //Выбираем таблицу(лист).
            Microsoft.Office.Interop.Excel.Worksheet ObjWorkSheet;
            ObjWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)ObjWorkBook.Sheets[1];
            string infoA = "";
            string infoB = "";
            //Выбираем первые сто записей из столбца.
            //Выбираем область таблицы. (в нашем случае просто ячейку)
<<<<<<< HEAD
            Microsoft.Office.Interop.Excel.Range rangeA = ObjWorkSheet.get_Range("A1");
            Microsoft.Office.Interop.Excel.Range rangeB = ObjWorkSheet.get_Range("B1");
            //Добавляем полученный из ячейки текст.
            infoA = rangeA.Text.ToString();
            infoB = rangeB.Text.ToString();
            Organization a = new Organization();
            a.FirstName = infoA;
            a.LastName = infoB;
            custdata.Add(a);
=======
>>>>>>> 7aba9fad65ebb305604ac0ae921a105771184c52
        }


    }
}
