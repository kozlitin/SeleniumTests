using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium;
using ActiveUp.Net.Mail;
using System.Text.RegularExpressions;
using System.Net;
using System.Net.Mail;
using System.Configuration;
using System.Collections.Specialized;

namespace SeleniumTests
{
    class Program
    {
        static void Main(string[] args)
        {

            ///////////////////// Читаем файл конфигурации /////////////////////

            string my_denovo_ua_username = ConfigurationManager.AppSettings.Get("my_denovo_ua_username");
            string my_denovo_ua_password = ConfigurationManager.AppSettings.Get("my_denovo_ua_password");

            string gmail_username = ConfigurationManager.AppSettings.Get("gmail_username");
            string gmail_password = ConfigurationManager.AppSettings.Get("gmail_password");
            
            string gmail_wait_seconds_string = ConfigurationManager.AppSettings.Get("gmail_wait_seconds");
            int gmail_wait_seconds = 15;
            try
            {
                gmail_wait_seconds = Int32.Parse(gmail_wait_seconds_string);
            }
            catch
            {
                gmail_wait_seconds = 15;
            }

            string gmail_attempts_number_string = ConfigurationManager.AppSettings.Get("gmail_attempts_number");
            int gmail_attempts_number = 3;
            try
            {
                gmail_attempts_number = Int32.Parse(gmail_attempts_number_string);
            }
            catch
            {
                gmail_attempts_number = 3;
            }
            if (gmail_attempts_number > 10) { 
                gmail_attempts_number = 10;
            }
            if (gmail_attempts_number < 1) { 
                gmail_attempts_number = 3;
            }

            string open_site_wait_seconds_string = ConfigurationManager.AppSettings.Get("open_site_wait_seconds");
            int open_site_wait_seconds = 15;
            try
            {
                open_site_wait_seconds = Int32.Parse(open_site_wait_seconds_string);
            }
            catch
            {
                open_site_wait_seconds = 15;
            }

            string final_email_from = ConfigurationManager.AppSettings.Get("final_email_from");
            string final_email_to = ConfigurationManager.AppSettings.Get("final_email_to");

            ///////////////////////////////////////////////////

            ChromeOptions options = new ChromeOptions();
            options.AddArgument("--headless");

            IWebDriver driver = new ChromeDriver(options);
            //IWebDriver driver = new ChromeDriver();

            StringBuilder testLog = new StringBuilder();

            LogMessage(testLog, "Начианем с ссылки https://support.my.denovo.ua/");

            driver.Navigate().GoToUrl("https://support.my.denovo.ua/");
            driver.Manage().Window.Size = new System.Drawing.Size(1521, 860);

            LogMessage(testLog, "Текущий URL: " + driver.Url);

            driver.FindElement(By.Id("inputUsername")).SendKeys(my_denovo_ua_username);
            driver.FindElement(By.Id("inputPassword")).Click();
            driver.FindElement(By.Id("inputPassword")).SendKeys(my_denovo_ua_password);
            driver.FindElement(By.CssSelector(".btn")).Click();

            LogMessage(testLog, "Пытаемся получить код авторизации ...");

            string passCode = "";
            int attemptNumber = 1;
            while(string.IsNullOrEmpty(passCode) && attemptNumber <= gmail_attempts_number)
            {
                Thread.Sleep(gmail_wait_seconds * 1000);
                LogMessage(testLog, "Попытка " + attemptNumber.ToString());
                passCode = GetAuthCodeFromGmail(gmail_username, gmail_password);
                attemptNumber++;
            }

            LogMessage(testLog, "Результирующий код авторизации: " + passCode);

            if (string.IsNullOrEmpty(passCode)) {
                driver.Quit();
                return;
            }

            driver.FindElement(By.Id("inputCode")).SendKeys(passCode);
            driver.FindElement(By.CssSelector(".btn")).Click();

            Thread.Sleep(open_site_wait_seconds * 1000);

            LogMessage(testLog, "Текущий URL: " + driver.Url);

            try {
                var elements = driver.FindElements(By.Id("MainHeaderSchemaCaptionValueLabel"));
                LogMessage(testLog, "MainHeaderSchemaCaptionValueLabel elements.Count = " + elements.Count.ToString());
                LogMessage(testLog, driver.FindElement(By.Id("MainHeaderSchemaCaptionValueLabel")).Text);
            }
            catch(OpenQA.Selenium.NoSuchElementException)
            {
                LogMessage(testLog, "Не найден элемент по id MainHeaderSchemaCaptionValueLabel");
            }

            LogMessage(testLog, "Переходим в личный кабинет ...");

            driver.FindElement(By.CssSelector(".t-btn-style-green")).Click();

            Thread.Sleep(open_site_wait_seconds * 1000);

            LogMessage(testLog, "Текущий URL: " + driver.Url);

            try
            {
                var elements = driver.FindElements(By.LinkText("Створити заявку"));
                LogMessage(testLog, "Створити заявку elements.Count = " + elements.Count.ToString());
            }
            catch (OpenQA.Selenium.NoSuchElementException)
            {
                LogMessage(testLog, "Не найден элемент Створити заявку");
            }

            LogMessage(testLog, "Переходим на сайт документации ...");

            driver.FindElement(By.CssSelector(".header_item:nth-child(1) span")).Click();

            Thread.Sleep(open_site_wait_seconds * 1000);

            driver.SwitchTo().Window(driver.WindowHandles.Last());

            LogMessage(testLog, "Текущий URL: " + driver.Url);

            try
            {
                var elements = driver.FindElements(By.CssSelector(".face_title"));
                LogMessage(testLog, ".face_title elements.Count = " + elements.Count.ToString());
                LogMessage(testLog, driver.FindElement(By.CssSelector(".face_title")).Text);
            }
            catch (OpenQA.Selenium.NoSuchElementException)
            {
                LogMessage(testLog, "Не найден элемент .face_title");
            }

            driver.Close();

            LogMessage(testLog, "Возвращаемся в личный кабинет ...");

            driver.SwitchTo().Window(driver.WindowHandles.First());

            LogMessage(testLog, "Текущий URL: " + driver.Url);

            LogMessage(testLog, "Выходим из личного кабинета ...");

            driver.FindElement(By.CssSelector(".header_item svg")).Click();

            Thread.Sleep(open_site_wait_seconds * 1000);

            LogMessage(testLog, "Текущий URL: " + driver.Url);

            try
            {
                var elements = driver.FindElements(By.Id("inputUsername"));
                LogMessage(testLog, "inputUsername elements.Count = " + elements.Count.ToString());
            }
            catch (OpenQA.Selenium.NoSuchElementException)
            {
                LogMessage(testLog, "Не найден элемент inputUsername");
            }

            LogMessage(testLog, "Тест окончен!");

            SendMailMessage(final_email_from, final_email_to, "Тест окончен!", testLog.ToString());

            driver.Quit();

        }

        static string GetAuthCodeFromGmail(string gmail_username, string gmail_password)
        {

            var mailRepository = new MailRepository(
                                        "imap.gmail.com",
                                        993,
                                        true,
                                        gmail_username,
                                        gmail_password
                                    );

            var emailList = mailRepository.GetUnreadMails("inbox");

            foreach (Message email in emailList)
            {
                if (email.From.ToString().Contains("Auth@de-novo.biz") && email.Subject.Contains("Підтвердження авторизації"))
                {
                    string pattern = @"(\d{4})";
                    Match m = Regex.Match(email.Subject, pattern, RegexOptions.IgnoreCase);
                    if (m.Success)
                        return m.Value;
                }
            }

            return "";

        }

        public static void SendMailMessage(string final_email_from, string final_email_to, string subject, string messageBody)
        {
            if (string.IsNullOrEmpty(final_email_to)) { 
                return; 
            }
            
            MailMessage message = new MailMessage(final_email_from, final_email_to);
            message.Subject = subject;
            message.Body = messageBody;
            System.Net.Mail.SmtpClient client = new System.Net.Mail.SmtpClient("mail.de-novo.biz", 10025);
            client.UseDefaultCredentials = true;

            try
            {
                client.Send(message);
            }
            catch (Exception ex)
            {
                Console.WriteLine("Exception caught in SendMailMessage(): {0}",
                    ex.ToString());
            }
        }

        static void LogMessage(StringBuilder log, string message)
        {
            log.AppendLine(message);
            Console.WriteLine(message);
        }

    }

    public class MailRepository
    {
        private Imap4Client client;

        public MailRepository(string mailServer, int port, bool ssl, string login, string password)
        {
            if (ssl)
                Client.ConnectSsl(mailServer, port);
            else
                Client.Connect(mailServer, port);
            Client.Login(login, password);
        }

        public IEnumerable<Message> GetAllMails(string mailBox)
        {
            return GetMails(mailBox, "ALL").Cast<Message>();
        }

        public IEnumerable<Message> GetUnreadMails(string mailBox)
        {
            return GetMails(mailBox, "UNSEEN").Cast<Message>();
        }

        protected Imap4Client Client
        {
            get { return client ?? (client = new Imap4Client()); }
        }

        private MessageCollection GetMails(string mailBox, string searchPhrase)
        {
            Mailbox mails = Client.SelectMailbox(mailBox);
            MessageCollection messages = mails.SearchParse(searchPhrase);
            return messages;
        }
    }

}
