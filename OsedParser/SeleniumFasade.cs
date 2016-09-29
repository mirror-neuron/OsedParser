using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Text.RegularExpressions;
using OpenQA.Selenium;

namespace OsedParser
{
    //автоматический выбор тулзы для скачивания
    public enum FileDownloadingTool
    {
        Selenium,
        WebClient
    }

    //режим захвата ссылок: по всем страницам или только первой (первые 30 карт)
    public enum LinkGetterPageState
    {
        FirstPage,
        AllPage
    }

    //режим захвата ссылок: забирает до первого сохраненного документа или всё подряд обновляет
    //
    public enum LinkGetterSyncState
    {
        AllLinks,
        OnlyNewLinks
    }

    class SeleniumFasade : IDisposable
    {
        private IWebDriver driver;
        private SQL sql;
        private string dnsId;
        private List<int> cardLinks = null;
        private FileDownloadingTool fileDownloadingTool;

        public SeleniumFasade(IWebDriver driver, SQL sql)
        {
            this.driver = driver;
            this.sql = sql;
        }

        /// <summary>
        /// Логин на сайт
        /// </summary>
        public void Login(string password)
        {
            try
            {
                //Открытие формы входа
                //
                driver.Navigate().GoToUrl(Program.baseURL);
                dnsId = Regex.Match(driver.Url, @"(?<=DNSID=)(\w*)").Value;

                //если ssl устарела, то используем для скачивания селениум, иначе вебклиент
                try
                {
                    driver.FindElement(By.XPath(".//*[@id='errorCode']"));
                    fileDownloadingTool = FileDownloadingTool.Selenium;
                }
                catch (NoSuchElementException)
                {
                    fileDownloadingTool = FileDownloadingTool.WebClient;
                }

                //ввод логина/пароля
                //
                driver.Navigate().GoToUrl(String.Format("{0}auth.php?DNSID={1}&group_id=18390", Program.baseURL, dnsId));
                driver.FindElement(By.XPath("//*[@id='user_id_chosen']/a/span")).Click();
                driver.FindElement(By.XPath("//li[text()[normalize-space(.)='Гудков И.Э.']]")).Click();
                driver.FindElement(By.Id("password")).SendKeys(password);

                //вход
                //
                Console.WriteLine("Login into account...");
                driver.FindElement(By.CssSelector("input[type=\"submit\"]")).Click();
                driver.WaitForPageLoad(10);

                //вошли?
                try
                {
                    driver.FindElement(By.XPath(".//*[@id='login_form']/table/tbody/tr[6]/td[2]/div/div[2]/span"));
                    sql.WriteLog("ПАРОЛЬ ИЗМЕНИЛСЯ! Откройте планировщик заданий и настройте запуск задания \"osed\" с параметором \"--pass <gudkov_password>\"");
                }
                catch (NoSuchElementException)
                { }
            }
            catch (Exception ex)
            {
                //логирование
                Logger.WriteToBase(ex);
            }
        }

        /// <summary>
        /// Отдает ссылки на карты путем перехода на все страницы
        /// </summary>
        public void GetCardLinks(LinkGetterPageState parserState = LinkGetterPageState.AllPage, LinkGetterSyncState syncState = LinkGetterSyncState.AllLinks)
        {
            //  sql.DeleteAll(); //грохнуть таблы
            cardLinks = new List<int>();
            int pageNumb = 1;

            try
            {
                ReadOnlyCollection<IWebElement> tableElements;

                bool getNextLink = true;

                while (getNextLink)
                {
                    //переход на следующую страницу
                    //
                    Console.WriteLine("Go to page {0}...", pageNumb);
                    driver.Navigate().GoToUrl(String.Format("{0}?type=0&DNSID={1}#page-{2}", Program.baseURL, dnsId, pageNumb));

                    //загрузилось?
                    driver.WaitForElementLoad(5, By.XPath("//*[@id='mtable']/tbody/tr[1]/td[4]"));
                    if (pageNumb > 1)
                    {
                        driver.Navigate().Refresh();
                        try
                        {
                            driver.WaitForElementLoad(5, By.XPath("//*[@id='mtable']/tbody/tr[1]/td[4]"));
                        }
                        catch (NoSuchElementException)
                        {  }
                    }

                    //забор ячеек
                    //
                    try
                    {
                        tableElements =
                            driver.FindElements(By.XPath("//*[@id='mtable']/tbody/tr[not(contains(@class, 'b'))]"));
                    }
                    catch (NoSuchElementException)
                    {
                        tableElements = null;
                    }

                    //последняя станица?
                    //
                    try
                    {
                        driver.FindElement(By.XPath("//*[@id='mtable']/tbody/tr/td[text()[normalize-space(.)='Нет записей.']]"));
                        getNextLink = false;
                        tableElements = null;
                    }
                    catch (NoSuchElementException)
                    { }

                    //забор ссылок
                    //
                    if (tableElements != null)
                    {
                        foreach (var element in tableElements)
                        {
                            try
                            {
                                int link =
                                    Convert.ToInt32(Regex.Match(element.FindElement(By.XPath("./td[4]/a")).GetAttribute("href"),
                                        "(?<=id=)([0-9]*)(?=&)").Value);
                                if (!sql.CardLinkExist(link))
                                {
                                    cardLinks.Add(link);
                                }
                                else if(syncState == LinkGetterSyncState.OnlyNewLinks)
                                {
                                    getNextLink = false;
                                }
                            }
                            catch (NoSuchElementException)
                            { }
                            if (!getNextLink)
                                break;
                        }
                    }

                    if (!getNextLink || parserState == LinkGetterPageState.FirstPage)
                    {
                        break;
                    }

                    pageNumb++;
                }
            }
            catch (Exception ex)
            {
                //логирование
                Logger.WriteToBase(ex);
            }
            cardLinks.Reverse();                            //парсим сначало самые старые доки
            cardLinks.AddRange(sql.getNotSynchronized());   //а потом пробуем ошибочные
        }

        public void ParseCards()
        {
            foreach (var cardLink in cardLinks)
            {
                Console.WriteLine("enter into card {0}...", cardLink);

                //вход
                //
                driver.Navigate()
                    .GoToUrl(String.Format("{0}document.card.php?id={1}&DNSID={2}", Program.baseURL, cardLink, dnsId));
                driver.WaitForPageLoad(10);

                //парсинг общих полей
                //
                int cl = Convert.ToInt32(cardLink);
                var card = new Card(sql.FindCardGuid(cl));
                card.LinkDocId = cl;
                try
                {
                    card.CardType = Card.DetectType(driver.FindElement(By.XPath("//*[@id='content']/h1")).Text);
                }
                catch (NoSuchElementException)
                { }
                try
                {
                    card.IncomeNumber = driver.FindElement(By.XPath("//*[@id='maintable']/tbody/tr[1]/td[2]")).Text;
                }
                catch (NoSuchElementException)
                { }
                try
                {
                    card.DocDate = driver.FindElement(By.XPath("//*[@id='maintable']/tbody/tr[1]/td[4]")).Text;
                        //driver.FindElement(By.XPath("//acronym[text()[normalize-space(.)='Дата документа:']]/../following-sibling::td[1]")).Text;
                }
                catch (NoSuchElementException)
                { }
                try
                {
                    card.RegDate = (card.CardType == "ВХ")
                    ? driver.FindElement(By.XPath("//*[@id='maintable']/tbody/tr[1]/td[4]")).Text
                    : card.DocDate;
                }
                catch (NoSuchElementException)
                { }


                //парсинг № документа
                //
                try
                {
                    if (card.CardType == "ВХ")
                        card.OutcomeNumber =
                            driver.FindElement(By.XPath(
                                "//acronym[text()[normalize-space(.)='№ документа:']]/../following-sibling::td[1]"))
                                .Text;
                }
                catch (NoSuchElementException)
                { }

                //парсинг Кому
                //
                try
                {
                    if (card.CardType == "ВХ")
                    {
                        card.ToName = driver
                            .FindElement(By.XPath("//*[@id='maintable']/tbody/tr[2]/td[2]")).Text;
                    }
                    else
                    {
                        card.ToName = driver
                            .FindElement(By.XPath("//*[@id='maintable']/tbody/tr[8]/td[2]")).Text;                  //not all information (example: id=1539065)
                    }
                }
                catch (NoSuchElementException)
                { }

                //парсинг От кого
                try
                {
                    if (card.CardType == "ВХ")
                    {
                        card.FromName =
                            driver.FindElement(By.XPath("//*[@id='maintable']/tbody/tr[7]/td[2]")).Text;
                    }
                }
                catch (NoSuchElementException)
                { }

                //парсинг Подпись
                try
                {
                    if (card.CardType != "ВХ")
                    {
                        card.SignName = driver
                            .FindElement(By.XPath("//*[@id='maintable']/tbody/tr[2]/td[2]")).Text;
                    }
                }
                catch (NoSuchElementException)
                { }

                //парсинг Исполнитель
                //
                try
                {
                    card.PerformName = driver.FindElement(By.XPath("//*[@id='td_prepaired_by']")).Text;
                }
                catch (NoSuchElementException)
                { }

                //парсинг Краткого содержания
                //
                try
                {
                    card.Summary =
                        driver.FindElement(
                            By.XPath(
                                "//acronym[text()[normalize-space(.)='Краткое содержание:']]/../following-sibling::td[1]"))
                            .Text;
                }
                catch (NoSuchElementException)
                {
                }

                //парсинг На №
                //
                card.References = new List<Tuple<bool, string, string, string>>();
                try
                {
                    driver.FindElement(By.XPath("//*[@id='action_sort_in_number']/div")).Click();
                    driver.WaitForElementLoad(5, By.XPath("//*[@id='action_sort_in_number']/table/tbody/tr"));
                    foreach (
                        var element in
                            driver.FindElements(By.XPath("//*[@id='action_sort_in_number']/table/tbody/tr")))
                    {
                        var temp = new Tuple<bool, string, string, string>
                            (true, element.FindElement(By.XPath("./td[1]")).Text,
                                element.FindElement(By.XPath("./td[2]")).Text,
                                element.FindElement(By.XPath("./td[3]")).Text);
                        card.References.Add(temp);
                    }
                }
                catch (NoSuchElementException)
                {}

                //парсинг На документ ссылаются:
                //
                try
                {
                    driver.FindElement(By.XPath("//*[@id='action_in_document_link']/div")).Click();
                    driver.WaitForElementLoad(5, By.XPath("//*[@id='action_in_document_link']/table/tbody/tr"));
                    foreach (
                        var element in
                            driver.FindElements(By.XPath("//*[@id='action_in_document_link']/table/tbody/tr")))
                    {
                        var temp = new Tuple<bool, string, string, string>
                            (false, element.FindElement(By.XPath("./td[1]")).Text,
                                element.FindElement(By.XPath("./td[2]")).Text,
                                element.FindElement(By.XPath("./td[3]")).Text);
                        card.References.Add(temp);
                    }
                }
                catch (NoSuchElementException)
                {
                }

                //парсинг приложенных файлов
                //
                try
                {
                    int index = 1;
                    var currentElement = driver.FindElement(By.XPath("//*[@id='files-first']/td[2]/a"));
                    Console.WriteLine("save attachment {0}...", index);
                    var currentLink = currentElement.GetAttribute("href");
                    var currentId =
                        Convert.ToInt32(Regex.Match(currentElement.GetAttribute("href"), "(?<=id=)([0-9]*)(?=&)").Value);
                    string fileName;
                    if (fileDownloadingTool == FileDownloadingTool.WebClient)
                    {
                        fileName = String.Format("{0}-{1}-fileId-{2}{3}", card.LinkDocId, index, currentId, Regex.Match(currentElement.Text, "(.\\w*$)").Value);
                        driver.DownloadFile(currentLink, Program.tempStoragePath + fileName);
                    }
                    else
                    {
                        fileName = currentElement.Text;
                        currentElement.Click();
                    }

                    card.AddFile(sql.FindFileGuid(fileName), fileName);

                    while (true)
                    {
                        index++;
                        currentElement =
                            currentElement.FindElement(By.XPath("following::tr[1]/td[2]/a"));
                        Console.WriteLine("save attachment {0}...", index);
                        currentLink = currentElement.GetAttribute("href");
                        currentId =
                            Convert.ToInt32(Regex.Match(currentElement.GetAttribute("href"), "(?<=id=)([0-9]*)(?=&)").Value);
                        if (fileDownloadingTool == FileDownloadingTool.WebClient)
                        {
                            fileName = String.Format("{0}-{1}-fileId-{2}{3}", card.LinkDocId, index, currentId, Regex.Match(currentElement.Text, "(.\\w*$)").Value);
                            driver.DownloadFile(currentLink, Program.tempStoragePath + fileName);
                        }
                        else
                        {
                            fileName = currentElement.Text;
                            if (System.IO.File.Exists(Program.tempStoragePath + fileName))
                                System.IO.File.Delete(Program.tempStoragePath + fileName);
                            currentElement.Click();
                        }

                        card.AddFile(sql.FindFileGuid(fileName), fileName);
                    }
                }
                catch (NoSuchElementException)
                {
                }

                //парсинг html документа для печати
                //
                try
                {
                    Console.WriteLine("press b07 button...");
                    try
                    {
                        driver.FindElement(By.XPath("//*[@id='b07']")).Click();
                    }
                    catch (NoSuchElementException) { }
                    driver.WaitForElementLoad(10, By.XPath("//*[@class='popup_info']//a[text()[normalize-space(.)='Карточка документа']]"));

                    Console.WriteLine("save html file...");
                    if (fileDownloadingTool == FileDownloadingTool.WebClient)
                    {
                        var pageLink = driver.FindElement(By.XPath("//*[@class='popup_info']//a[text()[normalize-space(.)='Карточка документа']]")).GetAttribute("href");
                        card.Html = driver.DownloadString(pageLink).Replace("rel=\"stylesheet\" href=\"/", "rel=\"stylesheet\" href=\"" + Program.baseURL);
                    }
                    else
                    {
                        string mainPage = driver.CurrentWindowHandle;
                        driver.FindElement(By.XPath("//*[@class='popup_info']//a[text()[normalize-space(.)='Карточка документа']]")).Click();
                        string newPage = driver.WindowHandles[1];
                        string pageLink = driver.SwitchTo().Window(newPage).Url;
                        card.Html = driver.DownloadString(pageLink).Replace("rel=\"stylesheet\" href=\"/", "rel=\"stylesheet\" href=\"" + Program.baseURL);
                        driver.Close();
                        driver.SwitchTo().Window(mainPage);
                    }
                }
                catch (NoSuchElementException)
                { }
                
                //парсинг файла для печати
                //
                try
                {
                    Console.WriteLine("save print file...");
                    var currentName = String.Format("{0}-printdoc.pdf", card.LinkDocId);

                    if (fileDownloadingTool == FileDownloadingTool.WebClient)
                    {
                        var currentLink = string.Format("{0}getpdfopt_mintrans.php?id={1}&DNSID={2}&doc=1", Program.baseURL, card.LinkDocId, dnsId);
                        driver.DownloadFile(currentLink, Program.tempStoragePath + currentName);
                    }
                    else
                    {
                        if (System.IO.File.Exists(Program.tempStoragePath + currentName))
                            System.IO.File.Delete(Program.tempStoragePath + currentName);
                        driver.FindElement(By.XPath("//*[@class='popup_info']//a[text()[normalize-space(.)='Только документ']]")).Click();
                    }
                    card.AddFile(sql.FindFileGuid(currentName), currentName);
                }
                catch (NoSuchElementException)
                {
                }

                ////parse scans
                //try
                //{
                //    driver.FindElement(By.XPath("//div[@id='d_all_pages_l']/a/b")).Click();
                //    System.Threading.Thread.Sleep(200);

                //    int i = 0;
                //    while (true)
                //    {
                //        var imgElement = driver.FindElement(By.Id("dpage_" + i++));
                //        Console.WriteLine("save scan {0}...", i);
                //        var imgLink = imgElement.GetAttribute("src");
                //        var imgName = String.Format("{0}-scan-{1}.jpg", card.LinkDocId, i);
                //        card.Files.Add(tempStoragePath + imgName);
                //        driver.DownloadExistsFile(imgLink, tempStoragePath + imgName);
                //    }
                //}
                //catch (NoSuchElementException)
                //{ }

                Console.WriteLine("Writing card to base...");
                card.WriteToBase(sql);
            }
        }

        public void Dispose()
        {
            Console.WriteLine("Exit...");
            try
            {
                driver.FindElement(By.XPath(".//*[@id='n_header']/div/div[3]/a")).Click();
                driver.WaitForPageLoad(10);
            }
            catch (NoSuchElementException)
            { }
            driver.Quit();
        }
    }
}
