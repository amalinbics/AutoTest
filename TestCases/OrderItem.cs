using NUnit.Framework;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using System.Threading;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data.OleDb;
using System.IO;
using Dapper;
using OpenQA.Selenium.Support.UI;

namespace AutoTest.TestCases
{
    class OrderItem
    {
        IWebDriver driver;
        public OrderItem()
        {
            driver = new ChromeDriver();
        }
        [SetUp]
        public void Initialize()
        {
            driver.Navigate().GoToUrl("http://training.openspan.com/login");
            driver.Manage().Window.Maximize();
            Thread.Sleep(2000);
        }
        [Test]
        public void ExecuteTest()
        {
            IWebElement userName = driver.FindElement(By.Name("user_name"));
            IWebElement password = driver.FindElement(By.Name("user_pass"));
            IWebElement btnLogin = driver.FindElement(By.Id("login_button"));
            userName.SendKeys("admin");
            password.SendKeys("admin");
            btnLogin.Click();

            var fileName  = Path.Combine( Directory.GetCurrentDirectory(), "Data/OrderInformation.xlsx");
            var con = string.Format(@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source = {0}; Extended Properties=Excel 12.0;", fileName);

            try
            {
                using (var connection = new OleDbConnection(con))
                {
                    connection.Open();


                    var query = string.Format("select * from [Products$];");
                    var orderData = connection.Query<OrderData>(query);

                    var billingAddresQuery = $"select * from [ShippingDetails$];";
                    var billingAddress = connection.Query<BillingAddress>(billingAddresQuery).SingleOrDefault();

                    connection.Close();

                    foreach (var item in orderData)
                    {
                        IWebElement productType = driver.FindElement(By.Name("productType"));
                        IWebElement product = driver.FindElement(By.Name("productsList"));


                        productType.SendKeys(item.ProductType);
                        product.SendKeys(item.Product);

                        IWebElement btnView = driver.FindElement(By.Id("viewButton"));

                        btnView.Click();

                        IWebElement quantity = driver.FindElement(By.Name("product_quantity"));
                        var selectElement = new SelectElement(quantity);
                        selectElement.SelectByValue(item.Qty);


                        IWebElement btnOrder = driver.FindElement(By.Name("Order"));

                        btnOrder.Click();

                        IWebElement btnEditCart = driver.FindElement(By.Name("edit_your_cart"));
                        btnEditCart.Click();



                        IWebElement next = driver.FindElement(By.Id("next1_button"));
                        next.Click();

                        IWebElement firstName = driver.FindElement(By.Name("bfirst_name"));
                        IWebElement lastName = driver.FindElement(By.Name("blast_name"));
                        IWebElement address = driver.FindElement(By.Name("bstreet_address"));
                        IWebElement zipCode = driver.FindElement(By.Name("bzip_code"));
                        IWebElement areaCode = driver.FindElement(By.Name("barea_code"));
                        IWebElement phone = driver.FindElement(By.Name("bprimary_phone"));

                        IWebElement nextInOrderPage = driver.FindElement(By.Id("next2_button"));

                        firstName.SendKeys(billingAddress.FirstName);
                        lastName.SendKeys(billingAddress.LastName);
                        address.SendKeys(billingAddress.Address);
                        zipCode.SendKeys("123");
                        areaCode.SendKeys(billingAddress.AreaCode);
                        phone.SendKeys(billingAddress.PrimaryPhone);

                        IWebElement btnShipToBilling = driver.FindElement(By.Id("ship_to_bill"));
                        btnShipToBilling.Click();

                        nextInOrderPage.Click();




                        IWebElement btnBillMe = driver.FindElement(By.Id("bill_me"));
                        btnBillMe.Click();

                        IWebElement btnSubmit = driver.FindElement(By.Id("submit_button"));
                        btnSubmit.Click();

                        IWebElement orderId = driver.FindElement(By.TagName("div p:first-of-type"));

                        item.OrderId = orderId.Text;

                        connection.Open();

                        var updatequery = $"update [Products$] set OrderId='{item.OrderId}' where Key='{item.Key}';";
                        var result = connection.Execute(updatequery);
                        connection.Close();

                        driver.Navigate().GoToUrl("http://training.openspan.com/home");



                    }

                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }
        [TearDown]
        public void EndTest()
        {
            driver.Close();
        }
    }
}
