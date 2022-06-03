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
using OpenQA.Selenium;
using Excel = Microsoft.Office.Interop.Excel;

namespace AlfaTask1
{
    public partial class MainWindow : System.Windows.Window
    {
        public enum WriteMode { oneRow, allRows }

        private WriteMode writeMode = WriteMode.oneRow;

        private IWebDriver driver;

        private By inputSubmitButton = By.XPath("//input[@type='submit']");

        private By inputEmailField = By.XPath("//input[@ng-reflect-name='labelEmail']");
        private By inputFirstNameField = By.XPath("//input[@ng-reflect-name='labelFirstName']");
        private By inputLastNameField = By.XPath("//input[@ng-reflect-name='labelLastName']");
        private By inputPhoneField = By.XPath("//input[@ng-reflect-name='labelPhone']");
        private By inputAddressField = By.XPath("//input[@ng-reflect-name='labelAddress']");
        private By inputCompanyNameField = By.XPath("//input[@ng-reflect-name='labelCompanyName']");
        private By inputCompanyRoleField = By.XPath("//input[@ng-reflect-name='labelRole']");

        public MainWindow()
        {
            InitializeComponent();
            //driver = new OpenQA.Selenium.Chrome.ChromeDriver();
            //readExcel();
            DriverSetup();
            this.ResizeMode = ResizeMode.NoResize;
        }

        private void DriverSetup()
        {
            driver = new OpenQA.Selenium.Chrome.ChromeDriver();
            driver.Navigate().GoToUrl("https://rpachallenge.com");
        }

        private void FillFields(string email, string firstName, string lastName, string phone, string address, string company, string role)
        {
            var emailField = driver.FindElement(inputEmailField);
            emailField.SendKeys(email);

            var firstNameField = driver.FindElement(inputFirstNameField);
            firstNameField.SendKeys(firstName);

            var lastNameField = driver.FindElement(inputLastNameField);
            lastNameField.SendKeys(lastName);

            var phoneField = driver.FindElement(inputPhoneField);
            phoneField.SendKeys(phone);

            var addressField = driver.FindElement(inputAddressField);
            addressField.SendKeys(address);

            var companyField = driver.FindElement(inputCompanyNameField);
            companyField.SendKeys(company);

            var roleField = driver.FindElement(inputCompanyRoleField);
            roleField.SendKeys(role);
        }
        private void Submit()
        {
            var submitButton = driver.FindElement(inputSubmitButton);
            submitButton.Click();
        }
        //private void FillField(By locator, string value)
        //{
        //    driver.FindElement(locator).SendKeys(value);
        //}
        private void WriteFromExcel()
        {
            Excel.Application excelFile = new Excel.Application().Application;

            string path = "";
            var dialog = new Microsoft.Win32.OpenFileDialog();
            dialog.Filter = "Лист Microsoft Excel|*.xls;*.xlsx;*.xlsm|All files (*.*)|*.*";

            bool? result = dialog.ShowDialog();

            if (result == true)
            {
                path = dialog.FileName;
            }

            Excel.Workbook workbook = excelFile.Workbooks.Open(path);
            Excel.Worksheet worksheet = workbook.Worksheets[1];

            Excel.Range range = worksheet.Rows[1];

            try
            {
                string fNameIndex = range.Find("First Name", Type.Missing, Excel.XlFindLookIn.xlValues, Excel.XlLookAt.xlPart).Address;
                fNameIndex = fNameIndex.Split('$')[1];
                string sNameIndex = range.Find("Last Name", Type.Missing, Excel.XlFindLookIn.xlValues, Excel.XlLookAt.xlPart).Address;
                sNameIndex = sNameIndex.Split('$')[1];
                string companyIndex = range.Find("Company Name", Type.Missing, Excel.XlFindLookIn.xlValues, Excel.XlLookAt.xlPart).Address;
                companyIndex = companyIndex.Split('$')[1];
                string roleIndex = range.Find("Role in Company", Type.Missing, Excel.XlFindLookIn.xlValues, Excel.XlLookAt.xlPart).Address;
                roleIndex = roleIndex.Split('$')[1];
                string emailIndex = range.Find("Email", Type.Missing, Excel.XlFindLookIn.xlValues, Excel.XlLookAt.xlPart).Address;
                emailIndex = emailIndex.Split('$')[1];
                string phoneIndex = range.Find("Phone", Type.Missing, Excel.XlFindLookIn.xlValues, Excel.XlLookAt.xlPart).Address;
                phoneIndex = phoneIndex.Split('$')[1];
                string addressIndex = range.Find("Address", Type.Missing, Excel.XlFindLookIn.xlValues, Excel.XlLookAt.xlPart).Address;
                addressIndex = addressIndex.Split('$')[1];

                range = worksheet.Columns[1];
                string lastIndexS = range.Find("", Type.Missing, Excel.XlFindLookIn.xlValues, Excel.XlLookAt.xlPart).Address;
                int lastIndex = Convert.ToInt32(lastIndexS.Split('$')[2]);

            
                for (int i = 2; i < lastIndex; i++)
                {
                    Excel.Range cell = worksheet.Range[fNameIndex+ i.ToString()];
                    string fNameValue = Convert.ToString(cell.Value);
                    cell = worksheet.Range[sNameIndex + i.ToString()];
                    string sNameValue = Convert.ToString(cell.Value);
                    cell = worksheet.Range[companyIndex + i.ToString()];
                    string company = Convert.ToString(cell.Value);
                    cell = worksheet.Range[roleIndex + i.ToString()];
                    string role = Convert.ToString(cell.Value);
                    cell = worksheet.Range[emailIndex + i.ToString()];
                    string email = Convert.ToString(cell.Value);
                    cell = worksheet.Range[phoneIndex + i.ToString()];
                    string phone = Convert.ToString(cell.Value);
                    cell = worksheet.Range[addressIndex + i.ToString()];
                    string address = Convert.ToString(cell.Value);

                    FillFields(
                        email,
                        fNameValue,
                        sNameValue,
                        phone,
                        address,
                        company,
                        role);
                    Submit();
                }
            }
            catch
            {
            
            }
        }

        private void writeButton_Click(object sender, RoutedEventArgs e)
        {
            switch (writeMode)
            {
                case WriteMode.oneRow:
                    FillFields(
                        emailTextBox.Text,
                        firstNameTextBox.Text,
                        lastNameTextBox.Text,
                        phoneTextBox.Text,
                        addressTextBox.Text,
                        companyTextBox.Text,
                        roleTextBox.Text);
                    break;
                case WriteMode.allRows:
                    WriteFromExcel();
                    break;
            }
        }

        private void writeAllFromExcelCheckBox_Checked(object sender, RoutedEventArgs e)
        {
            writeMode = WriteMode.allRows;
        }

        private void writeAllFromExcelCheckBox_Unchecked(object sender, RoutedEventArgs e)
        {
            writeMode = WriteMode.oneRow;
        }
    }
}
