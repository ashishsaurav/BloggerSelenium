using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Support.UI;
using System.Data;
using System.Diagnostics;
using Microsoft.Extensions.Configuration;
using ExcelDataReader;

public class Program
{
    public static IWebDriver driver { set; get; }
    public static DataTable dtblogger { get; set; }
    public static string profile { get; set; }
    public static string title { get; set; }
    public static IConfigurationRoot ConfigurationFile { get; set; }

    #region Main
    /// <summary>
    /// Main Function
    /// </summary>
    public static void Main()
    {
        try
        {
            #region WorkbookLoad and Initialize Variable
            Program program = new Program();
            ConfigurationFile = new ConfigurationBuilder()
                .SetBasePath(AppContext.BaseDirectory)
                .AddJsonFile("appsettings.json", optional: false, reloadOnChange: false)
                .Build();
            dtblogger = program.ReadBloggerkFile("autoblogger", dtblogger);
            dtblogger.Rows[0].Delete();
            dtblogger.AcceptChanges();
            title = ConfigurationFile.GetValue<string>("AppSettings:title");
            #endregion

            #region Add Urls of blogger
            Console.WriteLine("Please Enter Column to skip :- ");
            string colString = ConsoleReadline.ReadLine(10000);
            int colCount = Convert.ToInt32(colString);
            string rowString = "";
            int rowCount = 0;
            Console.WriteLine("Please hit enter after copy content and close all the browser:- ");
            ConsoleReadline.ReadLine(5000);
            List<string> bloggerUrls;
            #endregion

            while (true)
            {
                try
                {
                    #region Add Urls of blogger
                    Console.WriteLine("Please Enter Rows to skip :- ");

                    rowString = ConsoleReadline.ReadLine(10000);
                    rowCount = Convert.ToInt32(rowString);

                    bloggerUrls = dtblogger.AsEnumerable().Select(x => x.Field<string>(colCount.ToString())).ToList<string>();
                    if (string.IsNullOrWhiteSpace(bloggerUrls[0]))
                        break;
                    bloggerUrls.RemoveAll(s => string.IsNullOrWhiteSpace(s));
                    profile = bloggerUrls[0];
                    bloggerUrls.RemoveRange(0, rowCount + 1);
                    #endregion

                    #region Open Chrome Browser
                    driver = program.CreateBrowserDriver();
                    bool isNewPostPublished = false;
                    bool isNewPostOpen = false;
                    int count = 1;
                    int retryCount = 0;
                    #endregion

                    #region Process Each Row
                    foreach (var url in bloggerUrls)
                    {
                        if (string.IsNullOrWhiteSpace(url)) continue;
                        isNewPostPublished = false;
                        while (!isNewPostPublished)
                        {
                            driver.Navigate().GoToUrl(url);
                            if (driver.Url.Equals(url))
                            {
                                try
                                {
                                    #region Open New Post and Publish
                                    var newPostButtonXPath = "//*[@id=\"yDmH0d\"]/c-wiz/div[1]/gm-raised-drawer/div/div[2]/div[2]/c-wiz/div[2]/div/div";//"//*[@id=\"yDmH0d\"]/div[4]/div[2]/div/c-wiz/div[2]/div/div";
                                    var newPostButton = program.FindElement(By.XPath(newPostButtonXPath), 5);
                                    newPostButton.Click();
                                    Thread.Sleep(5000);
                                    isNewPostOpen = false;
                                    retryCount = 0;
                                    while (!isNewPostOpen)
                                    {
                                        if (driver.Url.Contains("edit"))
                                        {
                                            var newPostTitle = program.FindElement(By.CssSelector("input[aria-label='Title']"), 5);
                                            newPostTitle.SendKeys(title);

                                            var postBodyXPath = "//c-wiz[contains(@style,'visibility: visible')]/descendant::iframe[contains(@class,'ZW3ZFc')]";
                                            driver.SwitchTo().Frame(program.FindElement(By.XPath(postBodyXPath), 5));
                                            var newPostContent = program.FindElement(By.CssSelector("body.editable"), 5);
                                            newPostContent.SendKeys(Keys.Control + "v");
                                            driver.SwitchTo().DefaultContent();

                                            var publishXPath = "//c-wiz[contains(@style,'visibility: visible')]/descendant::div[contains(@aria-label,'Publish')]";
                                            var newPostPublish = program.FindElement(By.XPath(publishXPath), 5);
                                            newPostPublish.Click();

                                            var confirmXPath = "//div[contains(@class,'XfpsVe')]/descendant::div[contains(@data-id,'EBS5u')]";
                                            var newPostConfirm = program.FindElement(By.XPath(confirmXPath), 5);
                                            newPostConfirm.Click();

                                            while (!isNewPostPublished)
                                            {
                                                if (driver.Url.Equals(url))
                                                {
                                                    isNewPostPublished = true;
                                                    Console.WriteLine("Post Published - " + url + ", Count - " + count);
                                                    count++;
                                                }
                                            }
                                            isNewPostOpen = true;
                                        }
                                        else
                                        {
                                            if (retryCount > 1)
                                            {
                                                Console.WriteLine("Post Could not published - " + url + ", Count - " + count);
                                                isNewPostOpen = true;
                                                isNewPostPublished = true;
                                                count++;
                                            }
                                            else
                                            {
                                                retryCount++;
                                                Thread.Sleep(15000);
                                                if (!driver.Url.Contains("edit"))
                                                {
                                                    newPostButton.Click();
                                                    Thread.Sleep(5000);
                                                }

                                            }
                                        }
                                    }
                                    #endregion
                                }
                                catch (Exception ex)
                                {
                                    Console.WriteLine(ex.Message);
                                    Console.WriteLine("Post Could not published - " + url + ", Count - " + count);
                                    isNewPostPublished = true;
                                    isNewPostOpen = true;
                                    count++;
                                }
                            }
                            else
                            {
                                isNewPostPublished = true;
                                Console.WriteLine("Post info present - " + url + ", Count - " + count);
                                count++;
                            }
                        }
                    }
                    driver.Quit();

                    colCount++;
                    bloggerUrls.Clear();
                    #endregion
                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex.Message);
                    driver.Quit();
                }
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine(ex.Message);
            driver.Quit();
        }
    }
    #endregion

    #region Find Element
    /// <summary>
    /// Find Element
    /// </summary>
    /// <param name="by"></param>
    /// <param name="timeout"></param>
    /// <returns></returns>
    public IWebElement FindElement(By by, uint timeout)
    {
        try
        {
            if (string.IsNullOrWhiteSpace(by.Criteria))
                return null;
            var wait = new DefaultWait<IWebDriver>(driver);
            wait.Timeout = TimeSpan.FromSeconds(timeout);
            wait.IgnoreExceptionTypes(typeof(NoSuchElementException));
            return wait.Until(ctx =>
            {
                var elem = ctx.FindElement(by);
                if (!elem.Displayed && !elem.Enabled)
                    return null;

                return elem;
            });
        }
        catch (Exception ex)
        {
            Console.WriteLine(ex.Message);
            return null;
        }
    }
    #endregion

    #region Create Browser Driver
    /// <summary>
    /// Create Browser Driver
    /// </summary>
    /// <returns></returns>
    public IWebDriver CreateBrowserDriver()
    {
        try
        {
            var userdatadir = ConfigurationFile.GetValue<string>("AppSettings:userdatadir");
            var driverdir = ConfigurationFile.GetValue<string>("AppSettings:chromedir");

            var options = new ChromeOptions();
            options.AddArgument("--test-type");
            options.AddArgument("--ignore-certificate-errors");
            options.AddArgument("--no-sandbox");
            options.AddArgument("--disable-infobars");
            options.AddArgument("--start-maximized");
            options.AddArgument("--disable-dev-shm-usage");
            options.AddExcludedArgument("enable-automation");
            options.AddArgument(@"user-data-dir=" + userdatadir);
            options.AddArgument(@"profile-directory=" + profile);
            options.AcceptInsecureCertificates = true;

            return new ChromeDriver(driverdir, options, TimeSpan.FromSeconds(120));
        }
        catch (Exception ex)
        {
            throw;
        }
    }
    #endregion

    #region Read Blogger File
    /// <summary>
    /// Read Backlink File
    /// </summary>
    public DataTable ReadBloggerkFile(string bloggerFile, DataTable dtBloggerTable)
    {
        try
        {
            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
            var backlinkPath = ConfigurationFile.GetValue<string>("AppSettings:" + bloggerFile);

            FileStream stream = File.Open(backlinkPath, FileMode.Open, FileAccess.Read);
            IExcelDataReader excelReader = ExcelReaderFactory.CreateOpenXmlReader(stream);
            excelReader.AsDataSet(new ExcelDataSetConfiguration()
            {
                ConfigureDataTable = (_) => new ExcelDataTableConfiguration()
                {
                    UseHeaderRow = true
                }
            });

            DataSet result = excelReader.AsDataSet();

            dtBloggerTable = result.Tables[0];
            excelReader.Close();

            int i = 0;
            foreach (DataColumn column in dtBloggerTable.Columns)
            {
                column.ColumnName = i.ToString();
                i++;
            }

            dtBloggerTable.AcceptChanges();
            return dtBloggerTable;
        }
        catch (Exception ex)
        {
            throw ex;

        }
    }
    #endregion
}

#region Console Readline Class
/// <summary>
/// Console Readline Class
/// </summary>
class ConsoleReadline
{
    private static string inputLast;
    private static Thread inputThread = new Thread(inputThreadAction) { IsBackground = true };
    private static AutoResetEvent inputGet = new AutoResetEvent(false);
    private static AutoResetEvent inputGot = new AutoResetEvent(false);

    static ConsoleReadline()
    {
        inputThread.Start();
    }

    private static void inputThreadAction()
    {
        while (true)
        {
            inputGet.WaitOne();
            inputLast = Console.ReadLine();
            inputGot.Set();
        }
    }

    // omit the parameter to read a line without a timeout
    public static string ReadLine(int timeout = Timeout.Infinite)
    {
        if (timeout == Timeout.Infinite)
        {
            return Console.ReadLine();
        }
        else
        {
            var stopwatch = new Stopwatch();
            stopwatch.Start();

            while (stopwatch.ElapsedMilliseconds < timeout && !Console.KeyAvailable) ;

            if (Console.KeyAvailable)
            {
                inputGet.Set();
                inputGot.WaitOne();
                return inputLast;
            }
            else
            {
                return "0";
            }
        }
    }
}
#endregion
