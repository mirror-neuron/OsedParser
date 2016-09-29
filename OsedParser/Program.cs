using System;
using System.Reflection;
using OpenQA.Selenium.Chrome; //installed through NuGet

namespace OsedParser
{
    class Program
    {
        //текущий режим парсера и другие константы
        private static LinkGetterPageState linkGetterPageState = LinkGetterPageState.AllPage;
        private static LinkGetterSyncState linkGetterSyncState = LinkGetterSyncState.OnlyNewLinks;
        private static string defaultPassord = "";
        public const string baseURL = @"https://osed.mintrans.ru/";

        //различные настройки: для боевого сервера - бой, для остальных тест
        public static string tempStoragePath
        {
            get { return isServer() ? @"D:\Osed\Temp\" : @"D:\Sources\DV5\OsedParser\Temp\"; }
        }
        private static string dbName
        {
            get { return isServer() ? "docsbase5" : "docsvisionbase"; }
        }
        private static string dbCatalog
        {
            get { return isServer() ? "DV54" : "DV54_Test"; }
        }
        private static bool isServer()
        {
            return Environment.MachineName.ToUpper().Contains("DOCS") && Environment.MachineName.ToUpper().Contains("BASE");
        }

        static void Main(string[] args)
        {
            using (var sql = new SQL(dbName, dbCatalog))
            {
                try
                {
                    string pass = defaultPassord;
                    string help = "ЗАПУСКАЙТЕ ПРОГРАММУ С ПАРАМЕТРОМ!\nOsedStarter.bat файл или в планировщике параметр \"--pass <gudkov_password>\" у задания Osed";
                    if ((args.Length < 2 || args[0].Trim() != "--pass") && pass == "")
                    {
                        if (isServer())
                        {
                            sql.WriteLog(help);
                        }
                        else
                        {
                            Console.WriteLine(help);
                            System.Threading.Thread.Sleep(5000);
                        }
                        Environment.Exit(0);
                    }
                    else
                    {
                        pass = (pass == "") ? args[1].Trim() : pass;
                        Console.WriteLine("MachineName: {0}", Environment.MachineName.ToUpper());
                        Console.WriteLine("Passord: {0}", pass);

                        Console.WriteLine("============= START PARSE ===========");
                        Console.WriteLine("Waiting for open browser...");

                        var chromeOptions = new ChromeOptions();
                        chromeOptions.AddUserProfilePreference("download.default_directory", tempStoragePath);
                        chromeOptions.AddUserProfilePreference("disable-popup-blocking", "true");
                        using (var selenium = new SeleniumFasade(new ChromeDriver(chromeOptions), sql))
                        {
                            selenium.Login(pass);

                            selenium.GetCardLinks(linkGetterPageState, linkGetterSyncState);

                            selenium.ParseCards();
                        }

                        Console.WriteLine("============= END SUCCESSFULY PARSE ===========");
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine("============= PARSE FAIL ===========");
                    sql.WriteLog(ex.StackTrace);
                }
                finally
                {
                    Console.WriteLine("Activate ActionWorflow process...");
                    sql.activateAWProcess();
                    System.Threading.Thread.Sleep(5000);
                    if (linkGetterPageState == LinkGetterPageState.FirstPage)
                        Console.ReadKey();
                }
            }
        }
    }
}
