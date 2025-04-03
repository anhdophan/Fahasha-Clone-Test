using Microsoft.VisualStudio.TestTools.UnitTesting;
using NUnit.Framework;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Support.UI;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using Excel = Microsoft.Office.Interop.Excel;


namespace FahashaTest
{
    [TestFixture]
    public class Intergation_Test
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
            dataSheet = dataWorkBook.Sheets[4];
            xlRange = dataSheet.UsedRange;
        }
        
        [Test]
        public void Register_Order_Test()
        {
         
            for (int row = 6; row < 47; row+=21)
            {
              
               for (int step = 0; step < 21; step++)
               {
                    int currentRow = row + step;
                    //Đọc data
                    string stepAction = (xlRange.Cells[currentRow, 9] as Excel.Range).Value2?.ToString();
                    string testData = (xlRange.Cells[currentRow, 10] as Excel.Range).Value2?.ToString();
                   
                    if (testData == null) testData = "";
                    switch (stepAction.Trim())
                    {
                        case "Vào trang Đăng ký":
                            driver.Url = $"{testData}";
                            break;
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

                        case "Vào trang Đăng nhập":
                            driver.Url = $"{testData}";
                            break;

                        case "Nhập số điện thoại đăng nhập":
                            driver.FindElement(By.Id("login_username")).Clear();
                            driver.FindElement(By.Id("login_username")).SendKeys(testData);
                            break;

                        case "Nhập mật khẩu đăng nhập":
                            driver.FindElement(By.Id("login_password")).Clear();
                            driver.FindElement(By.Id("login_password")).SendKeys(testData);
                            driver.FindElement(By.Id("loginSubmit")).Click();
                            Thread.Sleep(2000);
                            break;

                        case "Tìm sản phẩm":
                            driver.FindElement(By.Name("_name")).Clear();
                            driver.FindElement(By.Name("_name")).SendKeys(testData);
                            driver.FindElement(By.Name("submitFind")).Click();
                            Thread.Sleep(2000);                          
                            break;
                        case "Nhấn thêm vào giỏ hàng":                           
                            driver.FindElement(By.Id("detailPro")).Click();
                            driver.FindElement(By.Id("addToCart")).Click();
                            break;

                        case "Thay đổi số lượng sản phẩm":
                            driver.FindElement(By.Name("cartQuantity")).Clear();
                            driver.FindElement(By.Name("cartQuantity")).SendKeys(testData);
                            driver.FindElement(By.Name("updateQuantity")).Click();
                            Thread.Sleep(2000);
                            break;

                        case "Nhấn thanh toán khi nhận hàng":                         
                            driver.FindElement(By.Id("COD")).Click();
                            Thread.Sleep(2000);
                            break;
                        case "Nhấn xác nhận đơn hàng":
                            driver.FindElement(By.Name("actionType")).Click();
                            Thread.Sleep(2000);
                            break;
                        case "Nhấn tiếp tục mua hàng":
                            driver.FindElement(By.Id("buyAgain")).Click();
                            break;
                        case "Vào trang quản lý đơn hàng":
                            driver.Url = $"{testData}";
                            break;
                        case "Tìm đơn hàng đặt muộn nhất":
                            driver.FindElements(By.Name("btnDetail")).LastOrDefault().Click();
                            break;
                        case "Chọn xem chi tiết":
                           
                            break;
                        case "Kiểm tra số lượng sản phẩm":
                            Thread.Sleep(1500);
                            //Lấy data expect result
                            string expectData = (xlRange.Cells[row, 11] as Excel.Range).Value2?.ToString();
                            //Lấy chuỗi string của thông báo
                            string countPro = driver.FindElement(By.Id("valueQuantity")).Text;
                            if (testData == countPro)
                            {
                                string actualValue = $"Số lượng sản phẩm của đơn hàng là {countPro}"; 

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
                                string actualValue = $"Số lượng sản phẩm của đơn hàng là {countPro}";
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

                            break;
                        case "Kết thúc":
                            driver.Url = "http://localhost:51529/LoginCustomer/ThongTinCaNhan";
                            driver.FindElement(By.Id("logOut")).Click();
                            Thread.Sleep(2000);
                            break;

                        default:                        
                            Console.WriteLine($"[WARNING] Step action không xác định: {stepAction}");
                            break;
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
