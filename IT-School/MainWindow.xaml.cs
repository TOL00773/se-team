using System.Collections.ObjectModel;
using System.IO;
using System.Net;
using System.Threading;
using System.Windows;
using Excel = Microsoft.Office.Interop.Excel;

namespace IT_School
{
    /// <summary>
    /// Логика взаимодействия для MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        ObservableCollection<Organization> custdata = new ObservableCollection<Organization>();

        public MainWindow()
        {
            InitializeComponent();
        }

        private void Button_Click_1(object sender, RoutedEventArgs e)
        {
            string filename = "data.xls";
            string adress = "http://www.admoblkaluga.ru/opendata/4027064263-ObrazovanieReestrUchrejd/data-124-structure-1.xls";
            WebClient myWebClient = new WebClient();
            myWebClient.DownloadFile(adress, filename);
        }
        private static string filename = Directory.GetCurrentDirectory() + @"\data.xls";


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
            object[,] arr = ObjWorkBook.get_Range("A1:B2000").Value;

<<<<<<< HEAD
            for (int i = 1; i < 501; i++)
            {
                Organization a = new Organization();
                
                Thread myThreadA = new Thread(func =>
                {
                    Excel.Range rangeA = ObjWorkSheet.get_Range("A" + i.ToString(), "A" + i.ToString());
                    info = rangeA.Text.ToString();
                    a.AccName = info;
                    Excel.Range range = ObjWorkSheet.get_Range("B" + i.ToString(), "B" + i.ToString());
                    info = range.Text.ToString();
                    a.Name = info;
                    Excel.Range rangeC = ObjWorkSheet.get_Range("C" + i.ToString(), "C" + i.ToString());
                    info = rangeC.Text.ToString();
                    a.Adress = info;

                });

                Thread myThreadB = new Thread(func =>
                {
                    myThreadA.Start(); //запускаем поток
                    Excel.Range rangeD = ObjWorkSheet.get_Range("D" + i.ToString(), "D" + i.ToString());
                    info = rangeD.Text.ToString();
                    a.GeoData = info;
                    Excel.Range rangeE = ObjWorkSheet.get_Range("E" + i.ToString(), "E" + i.ToString());
                    info = rangeE.Text.ToString();
                    a.WorkTime = info;
                    Excel.Range rangeF = ObjWorkSheet.get_Range("F" + i.ToString(), "F" + i.ToString());
                    info = rangeF.Text.ToString();
                    a.GosID = info;
                    Excel.Range rangeG = ObjWorkSheet.get_Range("G" + i.ToString(), "G" + i.ToString());
                    info = rangeG.Text.ToString();
                    a.Inn = info;
                });

                myThreadB.Start(); //запускаем поток

                Thread myThreadС = new Thread(func =>
                {
                    Excel.Range rangeH = ObjWorkSheet.get_Range("H" + i.ToString(), "H" + i.ToString());
                    info = rangeH.Text.ToString();
                    a.DateBegin = info;
                    Excel.Range rangeI = ObjWorkSheet.get_Range("I" + i.ToString(), "I" + i.ToString());
                    info = rangeI.Text.ToString();
                    a.GosAccReq = info;
                    Excel.Range rangeJ = ObjWorkSheet.get_Range("J" + i.ToString(), "J" + i.ToString());
                    info = rangeJ.Text.ToString();
                    a.DateExpire = info;
                    Excel.Range rangeK = ObjWorkSheet.get_Range("K" + i.ToString(), "K" + i.ToString());
                    info = rangeK.Text.ToString();
                    a.EduSpecs = info;
                });

                myThreadС.Start(); //запускаем поток

=======
            for (int i = 1; i < 201; i++)
            {
                Organization a = new Organization();
                Excel.Range rangeA = ObjWorkSheet.get_Range("A" + i.ToString(), "A" + i.ToString());
                info = rangeA.Text.ToString();
                a.AccName = info;
                Excel.Range range = ObjWorkSheet.get_Range("B" + i.ToString(), "B" + i.ToString());
                info = range.Text.ToString();
                a.Name = info;
                Excel.Range rangeC = ObjWorkSheet.get_Range("C" + i.ToString(), "C" + i.ToString());
                info = rangeC.Text.ToString();
                a.Adress = info;
                Excel.Range rangeD = ObjWorkSheet.get_Range("D" + i.ToString(), "D" + i.ToString());
                info = rangeD.Text.ToString();
                a.GeoData = info;
                Excel.Range rangeE = ObjWorkSheet.get_Range("E" + i.ToString(), "E" + i.ToString());
                info = rangeE.Text.ToString();
                a.WorkTime = info;
                Excel.Range rangeF = ObjWorkSheet.get_Range("F" + i.ToString(), "F" + i.ToString());
                info = rangeF.Text.ToString();
                a.GosID = info;
                Excel.Range rangeG = ObjWorkSheet.get_Range("G" + i.ToString(), "G" + i.ToString());
                info = rangeG.Text.ToString();
                a.Inn = info;
                Excel.Range rangeH = ObjWorkSheet.get_Range("H" + i.ToString(), "H" + i.ToString());
                info = rangeH.Text.ToString();
                a.DateBegin = info;
                Excel.Range rangeI = ObjWorkSheet.get_Range("I" + i.ToString(), "I" + i.ToString());
                info = rangeI.Text.ToString();
                a.GosAccReq = info;
                Excel.Range rangeJ = ObjWorkSheet.get_Range("J" + i.ToString(), "J" + i.ToString());
                info = rangeJ.Text.ToString();
                a.DateExpire = info;
                Excel.Range rangeK = ObjWorkSheet.get_Range("K" + i.ToString(), "K" + i.ToString());
                info = rangeK.Text.ToString();
                a.EduSpecs = info;
>>>>>>> 70c799d47748dc7db8e48343d6af9be25017aa37
                Excel.Range rangeL = ObjWorkSheet.get_Range("L" + i.ToString(), "L" + i.ToString());
                info = rangeL.Text.ToString();
                a.ReMake = info;
                Excel.Range rangeM = ObjWorkSheet.get_Range("M" + i.ToString(), "M" + i.ToString());
                info = rangeM.Text.ToString();
                a.StopStart = info;
                Excel.Range rangeN = ObjWorkSheet.get_Range("N" + i.ToString(), "N" + i.ToString());
                info = rangeN.Text.ToString();
                a.StopExec = info;
                Excel.Range rangeO = ObjWorkSheet.get_Range("O" + i.ToString(), "O" + i.ToString());
                info = rangeO.Text.ToString();
                a.Stop = info;

<<<<<<< HEAD
                if (a.AccName == "")
                {
                    break;
                }
=======
                custdata.Add(a);
>>>>>>> 70c799d47748dc7db8e48343d6af9be25017aa37

                custdata.Add(a);
                
            }

            DG1.ItemsSource = custdata;
            ObjExcel.Workbooks.Close();
        }

<<<<<<< HEAD
        private void Button_Click(object sender, RoutedEventArgs e)
        {

        }
=======
>>>>>>> 70c799d47748dc7db8e48343d6af9be25017aa37
    }
}