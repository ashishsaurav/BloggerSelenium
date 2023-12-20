using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Support.UI;
using System.Diagnostics;
using System.Data.OleDb;
using System.Data;
using System.Configuration;

public class Program
{
    public static IWebDriver driver { set; get; }
    public static string title { get; set; }
    public static string profile { get; set; }
    public static DataTable dtBlogger { get; set; }
    public static void Main()
    {
        try
        {
            #region WorkbookLoad and Initialize Variable
            Program program = new Program();
            program.ReadFile();
            var titlePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "bloggertitle.txt");
            title = File.ReadAllText(titlePath);
            Console.WriteLine("Please Enter Column to skip :- ");
            string colString = ConsoleReadline.ReadLine(10000);
            int colCount = Convert.ToInt32(colString) <=0 ? 0 : Convert.ToInt32(colString);
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
                    rowCount = Convert.ToInt32(rowString) <= 0 ? 0 : Convert.ToInt32(rowString);

                    bloggerUrls = dtBlogger.AsEnumerable().Select(x => x.Field<string>(colCount.ToString())).ToList<string>();
                    if (string.IsNullOrWhiteSpace(bloggerUrls[0]))
                        break;
                    bloggerUrls.RemoveAll(s => s == "");
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
                }
                catch(Exception ex)
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
    public IWebElement FindElement(By by, uint timeout)
    {
        try
        {
            var wait = new DefaultWait<IWebDriver>(driver);
            wait.Timeout = TimeSpan.FromSeconds(timeout);
            wait.IgnoreExceptionTypes(typeof(NoSuchElementException));
            return wait.Until(ctx =>
            {
                var elem = ctx.FindElement(by);
                if (!elem.Displayed)
                    return null;

                return elem;
            });
        }
        catch(Exception ex)
        {
            Console.WriteLine(ex.Message);
            return null;
        }
    }
    public IWebDriver CreateBrowserDriver()
    {
        try
        {
            string userdatadir = ConfigurationManager.AppSettings["userdatadir"];
            string driverdir = ConfigurationManager.AppSettings["chromedir"];

            var options = new ChromeOptions();
            options.AddArgument("--test-type");
            options.AddArgument("--ignore-certificate-errors");
            options.AddArgument("--no-sandbox");
            options.AddArgument("--disable-infobars");
            options.AddArgument("--start-maximized");
            options.AddArgument("--disable-dev-shm-usage");
            options.AddArgument(@"user-data-dir=" + userdatadir);
            options.AddArgument(@"profile-directory=" + profile);

            var directory = @driverdir;
            return new ChromeDriver(directory, options);
        }
        catch(Exception ex)
        {
            throw;
        }
    }
    public OleDbConnection InitializeOledbConnection()
    {
        var bloggerPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "bloggers.xlsx");
        string connString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + bloggerPath + ";Extended Properties=\"Excel 12.0;HDR=Yes;IMEX=1\"";
        return new OleDbConnection(connString);
    }
    public void ReadFile()
    {
        try
        {
            var oledbConn = InitializeOledbConnection();
            DataTable schemaTable = new DataTable();
            var OledbCmd = new OleDbCommand();
            OledbCmd.Connection = oledbConn;
            oledbConn.Open();
            OledbCmd.CommandText = "Select * from [Sheet1$]";
            OleDbDataReader dr = OledbCmd.ExecuteReader();
            if (dr.HasRows)
            {
                dtBlogger = new DataTable();
                dtBlogger.Columns.Add("0", typeof(string));
                dtBlogger.Columns.Add("1", typeof(string));
                dtBlogger.Columns.Add("2", typeof(string));
                dtBlogger.Columns.Add("3", typeof(string));
                dtBlogger.Columns.Add("4", typeof(string));
                dtBlogger.Columns.Add("5", typeof(string));
                dtBlogger.Columns.Add("6", typeof(string));
                dtBlogger.Columns.Add("7", typeof(string));
                dtBlogger.Columns.Add("8", typeof(string));
                dtBlogger.Columns.Add("9", typeof(string));
                dtBlogger.Columns.Add("10", typeof(string));
                dtBlogger.Columns.Add("11", typeof(string));
                dtBlogger.Columns.Add("12", typeof(string));
                dtBlogger.Columns.Add("13", typeof(string));
                dtBlogger.Columns.Add("14", typeof(string));
                while (dr.Read())
                {
                    dtBlogger.Rows.Add(dr[0], dr[1], dr[2], dr[3], dr[4], dr[5], dr[6], dr[7], dr[8],
                        dr[9], dr[10], dr[11], dr[12], dr[13], dr[14]);
                }
            }
            dr.Close();

            oledbConn.Close();
        }
        catch (Exception ex)
        {
            throw ex;

        }
    }
}
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