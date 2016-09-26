using System;
using System.Collections.Generic;
using System.Linq;

namespace OsedParser
{
    /// <summary>
    /// Structure for cards
    /// </summary>
    public class Card
    {
        public Guid OsedId;
        public int LinkDocId;
        public string CardType;
        public string IncomeNumber;
        public string OutcomeNumber = "";
        public string DocDate = "";
        public string RegDate = "";
        public string ToName = "";
        public string FromName = "";
        public string SignName = "";
        public string PerformName = "";
        public string Summary = "";
        public string Html = "";
        private List<Tuple<Guid, string>> files;
        public List<Tuple<Guid, string>> Files { get { return files; } }
        public List<Tuple<bool, string, string, string>> References = null;

        /// <summary>
        /// Создает новую карту, если данная карта уже существует то новый OsedGuid не генерируется
        /// </summary>
        /// <param name="findedGuid">Найденый OsedGuid</param>
        public Card(Guid? findedGuid)
        {
            if (findedGuid == null)
                OsedId = Guid.NewGuid();
            else
                OsedId = (Guid)findedGuid;
            files = new List<Tuple<Guid, string>>();
        }

        public void AddFile(Guid? guid, string name)
        {
            Guid FileId;
            if (guid == null)
                FileId = Guid.NewGuid();
            else
                FileId = (Guid)guid;
            files.Add(new Tuple<Guid, string>(FileId, name));
        }

        public void Print()
        {
            Console.WriteLine("linkId = {0}", LinkDocId);
            Console.WriteLine("cardType = {0}", CardType);
            Console.WriteLine("incomeNumber = {0}", IncomeNumber);
            Console.WriteLine("outcomeNumber = {0}", OutcomeNumber);
            Console.WriteLine("docDate = {0}", DocDate);
            Console.WriteLine("regDate = {0}", RegDate);
            Console.WriteLine("toName = {0}", ToName);
            Console.WriteLine("fromName = {0}", FromName);
            Console.WriteLine("signName = {0}", SignName);
            Console.WriteLine("performName = {0}", PerformName);
            Console.WriteLine("summary = {0}", Summary);
        }

        public void WriteToBase(SQL sql)
        {
            sql.WriteCardToBase(this);
        }

        /// <summary>
        /// Определняет тип карты внутри ОСЕД, в ДВ тип всегда входящий
        /// </summary>
        /// <param name="type"></param>
        /// <returns></returns>
        public static string DetectType(string type)
        {
            switch (type)
            {
                case "Входящие документы : карточка регистрации документа":
                    return "ВХ";
                case "Организационно-распорядительные документы : карточка регистрации документа":
                    return "ОРД";
                case "Внутренние документы : карточка регистрации документа":
                    return "ВН";
                case "Исходящие документы : карточка регистрации документа":
                    return "ИСХ";
                default:
                    return type;
            }
        }
    }
}
