using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
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
using System.IO;

namespace UnitedWay_WPF
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        List<string> EntriesList = new List<string>();
        List<string> NewEntriesList = new List<string>();

        public MainWindow()
        {
            InitializeComponent();
            
        }

        private void btnOpen_Click(object sender, RoutedEventArgs e)
        {
           // OpenFileDialog openfile = new OpenFileDialog();
           // openfile.DefaultExt = ".xlsx";
           // openfile.Filter = "(.xlsx)|*.xlsx";

            //var browsefile = openfile.ShowDialog();

            string FileName = "D:\\MSILaptop\\work\\iattendoutput.xlsx";


            if (File.Exists(FileName))
                {
                txtFilePath.Text = FileName;

                Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();

                Microsoft.Office.Interop.Excel.Workbook excelBook = excelApp.Workbooks.Open(txtFilePath.Text.ToString(), 0, true, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
                Microsoft.Office.Interop.Excel.Worksheet excelSheet = (Microsoft.Office.Interop.Excel.Worksheet)excelBook.Worksheets.get_Item(1); ;
                Microsoft.Office.Interop.Excel.Range excelRange = excelSheet.UsedRange;

                string strCellData = "";
                int rowCnt = 0;
                int colCnt = 0;
                double tickets = 0;
                int ticketsNum = 0;

                for (rowCnt = 2; rowCnt <= excelRange.Rows.Count; rowCnt++)
                {
                    for (colCnt = 1; colCnt <= 6; colCnt++)
                    {
                        if ((colCnt == 1) & (rowCnt != 2))
                        {
                            tickets = (excelRange.Cells[rowCnt, colCnt] as Microsoft.Office.Interop.Excel.Range).Value2;
                            ticketsNum = Convert.ToInt32(tickets);
                        }
                            
                        if ((colCnt == 6) & (rowCnt != 2))
                        {
                            strCellData = (string)(excelRange.Cells[rowCnt, colCnt] as Microsoft.Office.Interop.Excel.Range).Value2;
                            if (ticketsNum > 1)
                            {
                                while (ticketsNum > 0)
                                {
                                    EntriesList.Add(strCellData);
                                    ticketsNum --;
                                }
                            }
                            else
                            {
                                EntriesList.Add(strCellData);
                            }                             
                        }
                        else
                        {
                            continue;
                        }                
                    }
                }
                int ID = 0;
                foreach (var line in EntriesList)
                {
                    EntryList.Text += ID + "\t" + line + "\n";
                    ID++;
                }
                

                excelBook.Close(true, null, null);
                excelApp.Quit();
            }
            else
            {
                //add what happens if file isn't found
            }
        }

        private void btnClose_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }

        private void FindWinnerBtn_Click(object sender, RoutedEventArgs e)
        {
            //have a function for counting the number of items in the list
            int numEntries = EntriesList.Count();
            WinnerBox.Text = "Number of Entries: " + numEntries; // just for now to double check everything works
            
            
            // deduct 1 from it, and put range from 0 to that number into the random function
            int maxNum = numEntries - 1;
            // get the random number from function
            int winner = FindRandomNum(maxNum);
            //find the person with that element and reveal winner
            WinnerBox.Text = "/nFirst Winner is: " + EntriesList[winner];
            // take out the winner's name from the entire list
            RemoveWinnerFromList(EntriesList[winner]);
            DisplayNewList();
        }

        private int FindRandomNum (int maxNum)
        {
            // Instantiate random number generator.  
            Random rand = new Random();
            int min = 0;
            return rand.Next(min, maxNum);   
            
        }

        private void RemoveWinnerFromList(string winner)
        {   
            foreach (var item in EntriesList)
            {
                
                if (item != winner)
                {
                    NewEntriesList.Add(item);
                }      
            }
            EntriesList.Clear();
            foreach (var item in NewEntriesList)
            {
                EntriesList.Add(item);
            }
            NewEntriesList.Clear();
        }

        private void DisplayNewList()
        {
            EntryList.Text = "";
            int ID = 0;
            for (int i = 0; i < EntriesList.Count; i++)
            {
                EntryList.Text += ID + "\t" + EntriesList[i] + "\n";
                ID++;
            }
        }
    }
}
