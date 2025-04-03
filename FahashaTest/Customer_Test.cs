using Microsoft.VisualStudio.TestTools.UnitTesting;
using NUnit.Framework;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Support.UI;
using System;
using System.Collections.Generic;
using System.Text;
using System.Threading;
using Excel = Microsoft.Office.Interop.Excel;


namespace FahashaTest
{
    [TestFixture]
    public class Customer_Test
    {
        IWebDriver driver;
        IWebElement element;
        Excel.Application dataApp;//Mo excel
        Excel.Workbook dataWorkBook; //mo file excel
        Excel.Worksheet dataSheet; //mo sheet
        Excel.Range xlRange;

        [SetUp]
        public void SetUp()
        {
            driver = new ChromeDriver();
            driver.Url = "http://localhost:51529/LoginCustomer/RegisterUser";
            driver.Navigate();
            driver.Manage().Window.Maximize();
            Thread.Sleep(5000);

            dataApp = new Excel.Application();
            dataWorkBook = dataApp.Workbooks.Open("D:\\DBCLPM_LT\\FahashaTest\\Testscenario_testcase.xlsx");
            dataSheet = dataWorkBook.Sheets[3];
            xlRange = dataSheet.UsedRange;
        }
        public void ClearForm()
        {
            driver.FindElement(By.Id("NameCus")).Clear();
            driver.FindElement(By.Id("PhoneCus")).Clear();
            driver.FindElement(By.Id("EmailCus")).Clear();
            driver.FindElement(By.Id("Address")).Clear();           
            driver.FindElement(By.Id("PasswordUser")).Clear();                      
            driver.FindElement(By.Id("ConfirmPass")).Clear();
        }

        [Test]
        public void Register_Test()
        {
         
            for (int row = 7; row < 104; row+=7)
            {
                ClearForm();

               for (int step = 0; step < 7; step++)
               {
                    int currentRow = row + step;
                    //Đọc data
                    string stepAction = (xlRange.Cells[currentRow, 9] as Excel.Range).Value2?.ToString();
                    string testData = (xlRange.Cells[currentRow, 10] as Excel.Range).Value2?.ToString();
                   
                    if (testData == null) testData = "";
                    switch (stepAction.Trim())
                    {
                        case "Nhập tên khách hàng":
                            driver.FindElement(By.Id("NameCus")).Clear();
                            driver.FindElement(By.Id("NameCus")).SendKeys(testData);
                            break;

                        case "Nhập số điện thoại":
                            driver.FindElement(By.Id("PhoneCus")).Clear();
                            driver.FindElement(By.Id("PhoneCus")).SendKeys(testData);
                            break;

                        case "Nhập email":
                            driver.FindElement(By.Id("EmailCus")).Clear();
                            driver.FindElement(By.Id("EmailCus")).SendKeys(testData);
                            break;

                        case "Nhập địa chỉ":
                            driver.FindElement(By.Id("Address")).Clear();
                            driver.FindElement(By.Id("Address")).SendKeys(testData);
                            break;

                        case "Nhập mật khẩu":
                            driver.FindElement(By.Id("PasswordUser")).Clear();
                            driver.FindElement(By.Id("PasswordUser")).SendKeys(testData);
                            break;

                        case "Nhập xác nhận mật khẩu":
                            driver.FindElement(By.Id("ConfirmPass")).Clear();
                            driver.FindElement(By.Id("ConfirmPass")).SendKeys(testData);
                            break;

                        case "Nhấn đăng ký":
                            driver.FindElement(By.Id("registerConfirm")).Click();                            
                            break;

                        default:                        
                            Console.WriteLine($"[WARNING] Step action không xác định: {stepAction}");
                            break;
                    }
               }
                Thread.Sleep(1500);
                //Lấy data expect result
                string expectData = (xlRange.Cells[row, 11] as Excel.Range).Value2?.ToString();
                //Lấy chuỗi string của thông báo
                bool isFind = driver.FindElements(By.Id("notiregCus")).Count > 0;
                if (isFind)
                {
                    string actualValue = driver.FindElement(By.Id("notiregCus")).Text; // Nếu null thì gán chuỗi rỗng
                                                                                      
                    xlRange.Cells[row, 12] = actualValue;
                    if (actualValue == expectData)
                    {
                        xlRange.Cells[row, 13] = "Pass";
                    }
                    else
                    {
                        xlRange.Cells[row, 13] = "Fail";
                    }
                }
                else
                {
                    string actualValue = "Không tìm thấy thông báo";
                    xlRange.Cells[row, 12] = actualValue;
                    if (actualValue == expectData)
                    {
                        xlRange.Cells[row, 13] = "Pass";
                    }
                    else
                    {
                        xlRange.Cells[row, 13] = "Fail";
                    }
                }
                Thread.Sleep(2000);
            }

        }

        [TearDown]
        public void TearDown()
        {
            dataWorkBook.Save();
            dataWorkBook.Close();
            dataApp.Quit();
            driver.Quit();
        }
    }
}
