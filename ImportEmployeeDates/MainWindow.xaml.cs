/* Title:           Import Employee Dates
 * Date:            10-18-18
 * Author:          Terry Holmes
 * 
 * Description:     This will update employees start dates */

using System;
using System.Collections.Generic;
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
using Excel = Microsoft.Office.Interop.Excel;
using DataValidationDLL;
using NewEmployeeDLL;
using NewEventLogDLL;

namespace ImportEmployeeDates
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        //setting the classes
        WPFMessagesClass TheMessagesClass = new WPFMessagesClass();
        DataValidationClass TheDataValidationClass = new DataValidationClass();
        EmployeeClass TheEmployeeClass = new EmployeeClass();
        EventLogClass TheEventLogClass = new EventLogClass();

        //setting up the data
        FindEmployeeByPayIDDataSet TheFindEmployeeByPayIDDataSet = new FindEmployeeByPayIDDataSet();
        ImportedEmployeesDataSet TheImportEmployeesDataSet = new ImportedEmployeesDataSet();

        public MainWindow()
        {
            InitializeComponent();
        }

        private void btnImportExcel_Click(object sender, RoutedEventArgs e)
        {
            Excel.Application xlDropOrder;
            Excel.Workbook xlDropBook;
            Excel.Worksheet xlDropSheet;
            Excel.Range range;

            int intColumnRange = 0;
            int intCounter;
            int intNumberOfRecords;
            string strPayID;
            int intPayID;
            string strStartDate;
            DateTime datStartDate;
            string strEndDate;
            DateTime datEndDate;
            bool blnFatalError;            

            try
            {
                TheImportEmployeesDataSet.employees.Rows.Clear();

                Microsoft.Win32.OpenFileDialog dlg = new Microsoft.Win32.OpenFileDialog();
                dlg.FileName = "Document"; // Default file name
                dlg.DefaultExt = ".xlsx"; // Default file extension
                dlg.Filter = "Excel (.xlsx)|*.xlsx"; // Filter files by extension

                // Show open file dialog box
                Nullable<bool> result = dlg.ShowDialog();

                // Process open file dialog box results
                if (result == true)
                {
                    // Open document
                    string filename = dlg.FileName;
                }

                PleaseWait PleaseWait = new PleaseWait();
                PleaseWait.Show();

                xlDropOrder = new Excel.Application();
                xlDropBook = xlDropOrder.Workbooks.Open(dlg.FileName, 0, true, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
                xlDropSheet = (Excel.Worksheet)xlDropOrder.Worksheets.get_Item(1);

                range = xlDropSheet.UsedRange;
                intNumberOfRecords = range.Rows.Count;
                intColumnRange = range.Columns.Count;

                for (intCounter = 1; intCounter <= intNumberOfRecords; intCounter++)
                {
                    strPayID = Convert.ToString((range.Cells[intCounter, 5] as Excel.Range).Value2).ToUpper();
                    strStartDate = Convert.ToString((range.Cells[intCounter, 3] as Excel.Range).Value2).ToUpper();
                    if ((range.Cells[intCounter, 4] as Excel.Range).Value2 == null)
                    {
                        strEndDate = "12/31/2999";

                        datEndDate = Convert.ToDateTime(strEndDate);
                    }
                    else
                    {
                        strEndDate = Convert.ToString((range.Cells[intCounter, 4] as Excel.Range).Value2).ToUpper();

                        datEndDate = DateTime.FromOADate(Convert.ToDouble(strEndDate));
                    }


                    intPayID = Convert.ToInt32(strPayID);
                    datStartDate = DateTime.FromOADate(Convert.ToDouble(strStartDate));                    

                    TheFindEmployeeByPayIDDataSet = TheEmployeeClass.FindEmployeeByPayID(intPayID);

                    ImportedEmployeesDataSet.employeesRow NewEmployeeRow = TheImportEmployeesDataSet.employees.NewemployeesRow();

                    NewEmployeeRow.EmployeeID = TheFindEmployeeByPayIDDataSet.FindEmployeeByPayID[0].EmployeeID;
                    NewEmployeeRow.FirstName = TheFindEmployeeByPayIDDataSet.FindEmployeeByPayID[0].FirstName;
                    NewEmployeeRow.EndDate = datEndDate;
                    NewEmployeeRow.LastName = TheFindEmployeeByPayIDDataSet.FindEmployeeByPayID[0].LastName;
                    NewEmployeeRow.PayID = intPayID;
                    NewEmployeeRow.StartDate = datStartDate;

                    TheImportEmployeesDataSet.employees.Rows.Add(NewEmployeeRow);

                }

                PleaseWait.Close();
                dgrResults.ItemsSource = TheImportEmployeesDataSet.employees;
            }
            catch (Exception Ex)
            {
                TheEventLogClass.InsertEventLogEntry(DateTime.Now, "Import Employee Dates // Import Excel Button " + Ex.Message);

                TheMessagesClass.ErrorMessage(Ex.ToString());
            }
        }

        private void btnProcess_Click(object sender, RoutedEventArgs e)
        {
            int intCounter;
            int intNumberOfRecords;
            bool blnFatalError;
            int intEmployeeID;
            DateTime datStartDate;
            DateTime datEndDate;

            try
            {
                intNumberOfRecords = TheImportEmployeesDataSet.employees.Rows.Count - 1;

                for(intCounter = 0; intCounter <= intNumberOfRecords; intCounter++)
                {
                    intEmployeeID = TheImportEmployeesDataSet.employees[intCounter].EmployeeID;
                    datStartDate = TheImportEmployeesDataSet.employees[intCounter].StartDate;
                    datEndDate = TheImportEmployeesDataSet.employees[intCounter].EndDate;

                    blnFatalError = TheEmployeeClass.UpdateEmployeeStartDate(intEmployeeID, datStartDate);

                    if (blnFatalError == true)
                        throw new Exception();

                    blnFatalError = TheEmployeeClass.UpdateEmployeeEndDate(intEmployeeID, datEndDate);

                    if (blnFatalError == true)
                        throw new Exception();
                }

                TheMessagesClass.InformationMessage("Employees Updated");
            }
            catch (Exception Ex)
            {
                TheEventLogClass.InsertEventLogEntry(DateTime.Now, "Import Employee Dates // Process Button " + Ex.Message);

                TheMessagesClass.ErrorMessage(Ex.ToString());
            }
        }
    }
}
