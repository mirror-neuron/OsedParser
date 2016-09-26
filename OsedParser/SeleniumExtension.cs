using System;
using System.Linq;
using OpenQA.Selenium;              //installed through NuGet
using OpenQA.Selenium.Support.UI;   //installed through NuGet
using System.IO;
using System.Drawing;
using System.Net;

namespace OsedParser
{
    public static class ExtensionToSelenium
    {
        public static IWebElement WaitForElementLoad(this IWebDriver driver, int maxWaitTimeInSeconds, By by)
        {
            for (int i = 0; i < maxWaitTimeInSeconds; i++)
            {
                try
                {
                    return driver.FindElement(by);
                }
                catch (NoSuchElementException)
                {
                    System.Threading.Thread.Sleep(100);
                }
            }
            throw new NoSuchElementException();
        }

        /// <summary>
        /// Waiting for page loading
        /// </summary>
        public static void WaitForPageLoad(this IWebDriver driver, int maxWaitTimeInSeconds)
        {
            string state = string.Empty;
            try
            {
                WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(maxWaitTimeInSeconds));

                //Checks every 500 ms whether predicate returns true if returns exit otherwise keep trying till it returns ture
                wait.Until(d => {

                    try
                    {
                        state = ((IJavaScriptExecutor)driver).ExecuteScript(@"return document.readyState").ToString();
                    }
                    catch (InvalidOperationException)
                    {
                        //Ignore
                    }
                    catch (NoSuchWindowException)
                    {
                        //when popup is closed, switch to last windows
                        driver.SwitchTo().Window(driver.WindowHandles.Last());
                    }
                    //In IE7 there are chances we may get state as loaded instead of complete
                    return (state.Equals("complete", StringComparison.InvariantCultureIgnoreCase) || state.Equals("loaded", StringComparison.InvariantCultureIgnoreCase));

                });
            }
            catch (TimeoutException)
            {
                //sometimes Page remains in Interactive mode and never becomes Complete, then we can still try to access the controls
                if (!state.Equals("interactive", StringComparison.InvariantCultureIgnoreCase))
                    throw;
            }
            catch (NullReferenceException)
            {
                //sometimes Page remains in Interactive mode and never becomes Complete, then we can still try to access the controls
                if (!state.Equals("interactive", StringComparison.InvariantCultureIgnoreCase))
                    throw;
            }
            catch (WebDriverException)
            {
                if (driver.WindowHandles.Count == 1)
                {
                    driver.SwitchTo().Window(driver.WindowHandles[0]);
                }
                state = ((IJavaScriptExecutor)driver).ExecuteScript(@"return document.readyState").ToString();
                if (!(state.Equals("complete", StringComparison.InvariantCultureIgnoreCase) || state.Equals("loaded", StringComparison.InvariantCultureIgnoreCase)))
                    throw;
            }
        }

        /// <summary>
        /// Taking screenshot of element by "By"
        /// </summary>
        public static Bitmap TakeScreenshot(this IWebDriver _driver, By by)
        {
            var screenshotDriver = _driver as ITakesScreenshot;
            Screenshot screenshot = screenshotDriver.GetScreenshot();

            var bmpScreen = new Bitmap(new MemoryStream(screenshot.AsByteArray));
            IWebElement element = _driver.FindElement(by);
            var cropArea = new Rectangle(element.Location, element.Size);
            return bmpScreen.Clone(cropArea, bmpScreen.PixelFormat);
        }

        /// <summary>
        /// Download file by url with cookie
        /// </summary>
        public static void DownloadFile(this IWebDriver driver, string url, string localPath)
        {
            try
            {
                using (var client = new WebClient())
                {
                    client.Headers[HttpRequestHeader.Cookie] = driver.CookieString();
                    if (Directory.Exists(Path.GetDirectoryName(localPath)))
                    {
                        client.DownloadFile(url, localPath);
                    }
                }
            }
            catch (Exception ex)
            {
                Logger.WriteToBase(ex);
            }
        }

        /// <summary>
        /// Download file by url with cookie
        /// </summary>
        public static string DownloadString(this IWebDriver driver, string url)
        {
            try
            {
                using (var client = new WebClient())
                {
                    client.Headers[HttpRequestHeader.Cookie] = driver.CookieString();
                    return client.DownloadString(url);
                }
            }
            catch (Exception ex)
            {
                Logger.WriteToBase(ex);
                throw new Exception();
            }
        }

        /// <summary>
        /// Return current cookie
        /// </summary>
        public  static string CookieString(this IWebDriver driver)
        {
            var cookies = driver.Manage().Cookies.AllCookies;
            return string.Join("; ", cookies.Select(c => string.Format("{0}={1}", c.Name, c.Value)));
        }
    }
}
