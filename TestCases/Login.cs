using NUnit.Framework;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace AutoTest.TestCases
{
    class Login
    {
        IWebDriver driver;
        public Login()
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

        }
        [TearDown]
        public void EndTest()
        {
            driver.Close();     
        }
    }
}
