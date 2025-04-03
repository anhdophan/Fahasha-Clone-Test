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
    public class Admin_Test
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
            driver.Url = "http://localhost:51529/Product/QuanlySP";
            driver.Navigate();
            driver.Manage().Window.Maximize();
            Thread.Sleep(5000);

            dataApp = new Excel.Application();
            dataWorkBook = dataApp.Workbooks.Open("D:\\DBCLPM_LT\\FahashaTest\\Testscenario_testcase.xlsx");
            dataSheet = dataWorkBook.Sheets[2];
            xlRange = dataSheet.UsedRange;
        }
        public void ClearForm()
        {
            driver.FindElement(By.Id("NamePro")).Clear();
            driver.FindElement(By.Id("Price")).Clear();
            driver.FindElement(By.Id("DecriptionSmall")).Clear();
            driver.FindElement(By.Id("DescriptionBig")).Clear();           
            driver.FindElement(By.Id("Author")).Clear();                      
            driver.FindElement(By.Id("Series")).Clear();
            driver.FindElement(By.Id("Publishingyear")).Clear();
            driver.FindElement(By.Id("C_Language")).Clear();
            driver.FindElement(By.Id("Quantity")).Clear();
        }

        [Test]
        public void Add_Product_Test()
        {
            for (int row = 7; row < 216; row+=14)
            {
                driver.FindElement(By.Id("addPro")).Click();
                Thread.Sleep(2500);
                ClearForm();

               for (int step = 0; step < 14; step++)
               {
                    int currentRow = row + step;
                    //Đọc data
                    string stepAction = (xlRange.Cells[currentRow, 9] as Excel.Range).Value2?.ToString();
                    string testData = (xlRange.Cells[currentRow, 10] as Excel.Range).Value2?.ToString();
                   
                    if (testData == null) testData = "";
                    switch (stepAction.Trim())
                    {
                        case "Nhập tên sản phẩm":
                            driver.FindElement(By.Id("NamePro")).Clear();
                            driver.FindElement(By.Id("NamePro")).SendKeys(testData);
                            break;

                        case "Nhập giá sản phẩm":
                            driver.FindElement(By.Id("Price")).Clear();
                            driver.FindElement(By.Id("Price")).SendKeys(testData);
                            break;

                        case "Nhập mô tả nhỏ":
                            driver.FindElement(By.Id("DecriptionSmall")).Clear();
                            driver.FindElement(By.Id("DecriptionSmall")).SendKeys(testData);
                            break;

                        case "Nhập mô tả chi tiết":
                            driver.FindElement(By.Id("DescriptionBig")).Clear();
                            driver.FindElement(By.Id("DescriptionBig")).SendKeys(testData);
                            break;

                        case "Chọn loại danh mục sản phẩm":
                            var cateDropdown = driver.FindElement(By.Id("Category"));
                            SelectElement selectCate = new SelectElement(cateDropdown);
                            selectCate.SelectByText(testData);                          
                            break;

                        case "Nhập tên tác giả":
                            driver.FindElement(By.Id("Author")).Clear();
                            driver.FindElement(By.Id("Author")).SendKeys(testData);
                            break;

                        case "Chọn nhà cung cấp":
                            var supplierDropdown = driver.FindElement(By.Id("Supplier"));
                            SelectElement selectSupplier = new SelectElement(supplierDropdown);
                            selectSupplier.SelectByText(testData);
                            
                            break;

                        case "Chọn nhà suất bản":
                            var publisherDropdown = driver.FindElement(By.Id("Publisher"));
                            SelectElement selectPublisher = new SelectElement(publisherDropdown);
                            selectPublisher.SelectByText(testData);                           
                            break;

                        case "Nhập series":
                            driver.FindElement(By.Id("Series")).Clear();
                            driver.FindElement(By.Id("Series")).SendKeys(testData);
                            break;

                        case "Nhập năm phát hành":
                            driver.FindElement(By.Id("Publishingyear")).Clear();
                            driver.FindElement(By.Id("Publishingyear")).SendKeys(testData);
                            break;

                        case "Nhập ngôn ngữ":
                            driver.FindElement(By.Id("C_Language")).Clear();
                            driver.FindElement(By.Id("C_Language")).SendKeys(testData);
                            break;

                        case "Nhập số lượng":
                            driver.FindElement(By.Id("Quantity")).Clear();
                            driver.FindElement(By.Id("Quantity")).SendKeys(testData);
                            break;

                        case "Chọn file ảnh":                           
                            driver.FindElement(By.Name("UploadImage")).SendKeys(@"E:\0502-G03 - web Bán Sách\DoAn\DoAn\Content\images\" + testData);                           
                            break;

                        case "Nhấn tạo mới":
                            driver.FindElement(By.Name("createPro")).Click();                            
                            break;

                        default:                        
                            Console.WriteLine($"[WARNING] Step action không xác định: {stepAction}");
                            break;
                    }
               }
                Thread.Sleep(1000);
                string expectData = (xlRange.Cells[row, 11] as Excel.Range).Value2?.ToString();
                bool isFind = driver.FindElements(By.Id("notiAddPro")).Count > 0;
                if (isFind)
                {
                    string actualValue = driver.FindElement(By.Id("notiAddPro")).Text; // Nếu null thì gán chuỗi rỗng

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
        [Test]
        public void Update_Product_Test()
        {            
            for (int row = 217; row < 398; row += 14)
            {

                
                    driver.FindElement(By.Name("searchTerm")).SendKeys("Số đỏ");
                    driver.FindElement(By.Name("searchTermBtn")).Click();
                    Thread.Sleep(2500);
                    driver.FindElement(By.Name("editPro")).Click();
                    ClearForm();

                    for (int step = 0; step < 14; step++)
                    {
                        int currentRow = row + step;
                        string stepAction = (xlRange.Cells[currentRow, 9] as Excel.Range).Value2?.ToString();
                        string testData = (xlRange.Cells[currentRow, 10] as Excel.Range).Value2?.ToString() ?? "";

                                                
                            switch (stepAction.Trim())
                            {
                                case "Nhập tên sản phẩm":
                                    driver.FindElement(By.Id("NamePro")).Clear();
                                    driver.FindElement(By.Id("NamePro")).SendKeys(testData);
                                    break;

                                case "Nhập giá sản phẩm":
                                    driver.FindElement(By.Id("Price")).Clear();
                                    driver.FindElement(By.Id("Price")).SendKeys(testData);
                                    break;

                                case "Nhập mô tả nhỏ":
                                    driver.FindElement(By.Id("DecriptionSmall")).Clear();
                                    driver.FindElement(By.Id("DecriptionSmall")).SendKeys(testData);
                                    break;

                                case "Nhập mô tả chi tiết":
                                    driver.FindElement(By.Id("DescriptionBig")).Clear();
                                    driver.FindElement(By.Id("DescriptionBig")).SendKeys(testData);
                                    break;

                                case "Chọn loại danh mục sản phẩm":
                                    var cateDropdown = driver.FindElement(By.Id("Category"));
                                    SelectElement selectCate = new SelectElement(cateDropdown);
                                    selectCate.SelectByText(testData);
                                    break;

                                case "Nhập tên tác giả":
                                    driver.FindElement(By.Id("Author")).Clear();
                                    driver.FindElement(By.Id("Author")).SendKeys(testData);
                                    break;

                                case "Chọn nhà cung cấp":
                                    var supplierDropdown = driver.FindElement(By.Id("Supplier"));
                                    SelectElement selectSupplier = new SelectElement(supplierDropdown);
                                    selectSupplier.SelectByText(testData);
                                    break;

                                case "Chọn nhà suất bản":
                                    var publisherDropdown = driver.FindElement(By.Id("Publisher"));
                                    SelectElement selectPublisher = new SelectElement(publisherDropdown);
                                    selectPublisher.SelectByText(testData);
                                    break;

                                case "Nhập series":
                                    driver.FindElement(By.Id("Series")).Clear();
                                    driver.FindElement(By.Id("Series")).SendKeys(testData);
                                    break;

                                case "Nhập năm phát hành":
                                    driver.FindElement(By.Id("Publishingyear")).Clear();
                                    driver.FindElement(By.Id("Publishingyear")).SendKeys(testData);
                                    break;

                                case "Nhập ngôn ngữ":
                                    driver.FindElement(By.Id("C_Language")).Clear();
                                    driver.FindElement(By.Id("C_Language")).SendKeys(testData);
                                    break;

                                case "Nhập số lượng":
                                    driver.FindElement(By.Id("Quantity")).Clear();
                                    driver.FindElement(By.Id("Quantity")).SendKeys(testData);
                                    break;

                                case "Chọn file ảnh":
                                    driver.FindElement(By.Name("UploadImage")).SendKeys(@"E:\\0502-G03 - web Bán Sách\\DoAn\\DoAn\\Content\\images\\" + testData);
                                    break;

                                case "Nhấn tạo mới":
                                    driver.FindElement(By.Name("updatePro")).Click();
                                    break;

                                default:
                                    Console.WriteLine($"Step action không xác định: {stepAction}");
                                    break;
                            }
                       
                    }
                    Thread.Sleep(1000);
                    string expectData = (xlRange.Cells[row, 11] as Excel.Range).Value2?.ToString();
                bool isFind = driver.FindElements(By.Id("notiAddPro")).Count > 0;
                if (isFind)
                {
                    string actualValue = driver.FindElement(By.Id("notiAddPro")).Text; // Nếu null thì gán chuỗi rỗng

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
                    Thread.Sleep (2000);
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
