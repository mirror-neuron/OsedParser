using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.IO;

namespace OsedParser
{
    public sealed class SQL : IDisposable
    {
        private SqlConnection sqlConnection;

        public SQL(string baseName, string catalog)
        {
            sqlConnection = new SqlConnection(
                    string.Format("Data Source={0};Initial Catalog={1};User Id=tunedv;Password=xfKgso6HL7;MultipleActiveResultSets=True;Connection Timeout=90;",
                        baseName, catalog));
            sqlConnection.Open();
            CheckConnection();
        }

        /// <summary>
        /// Проверяем не пропал ли коннект к базе и пытаемся восстановить
        /// </summary>
        private void CheckConnection()
        {
            try
            {
                if (sqlConnection == null && sqlConnection.State != ConnectionState.Open)
                {
                    sqlConnection.Open();
                }
                if (sqlConnection == null && sqlConnection.State != ConnectionState.Open)
                {
                    throw new Exception();
                }
            }
            catch (Exception ex)
            {
                //логирование
                Logger.WriteToBase(ex);
            }
        }

        /// <summary>
        /// Query command
        /// </summary>
        public DataTable PullDataTable(string command)
        {
            CheckConnection();
            var dataTable = new DataTable();
            using (SqlDataAdapter da = new SqlDataAdapter(command, sqlConnection))
            {
                da.Fill(dataTable);
            }
            return dataTable;
        }

        /// <summary>
        /// Execute nonQuery command
        /// </summary>
        public void ExecuteNonQuery(string command)
        {
            CheckConnection();
            (new SqlCommand(command, sqlConnection)).ExecuteNonQuery();
        }

        /// <summary>
        /// Кидаю активность 
        /// </summary>
        public void activateAWProcess()
        {
            CheckConnection();
            (new SqlCommand(@"INSERT INTO [dbo].[awf_ExtensionProcesses]
               ([RowID], [CardID], [OperationType], [ProcessType], [ProcessTime], [CreationTime])
                VALUES
               (NEWID(), NEWID(), 'CreateCardFromOsed', 'CreateCardFromOsed', GETDATE(), GETDATE());", sqlConnection)).ExecuteNonQuery();
        }

        /// <summary>
        /// Writing card with scan and attachmenst to base
        /// </summary>
        public void WriteCardToBase(Card card)
        {
            //try
            //{
            string cmd = string.Format(@" SET DATEFORMAT DMY;
                        MERGE dbo.[4doc_Osed] A
                            USING (VALUES ('{0}', {1}, 0, '{2}','{3}','{4}','{5}','{6}','{7}','{8}','{9}','{10}','{11}', '{12}')) B 
                            (OsedId, LinkDocId, SyncSuccess, CardType, IncomeNumber, OutcomeNumber,
                            DocDate, RegDate, Summary, ToName, FromName, SignName, PerformName, Html) ON (A.LinkDocId = B.LinkDocId)
                        WHEN MATCHED 
                            THEN UPDATE SET LinkDocId=B.LinkDocId, SyncSuccess=B.SyncSuccess, CardType=B.CardType,
                            IncomeNumber=B.IncomeNumber, OutcomeNumber=B.OutcomeNumber, DocDate=B.DocDate, RegDate=B.RegDate,
                            Summary=B.Summary, ToName=B.ToName, FromName=B.FromName, SignName=B.SignName, PerformName=B.PerformName, Html=B.Html
                        WHEN NOT MATCHED 
                            THEN INSERT (OsedId, LinkDocId, SyncSuccess, CardType, IncomeNumber, OutcomeNumber,
                            DocDate, RegDate, Summary, ToName, FromName, SignName, PerformName, Html)
                        VALUES (B.OsedId, B.LinkDocId, B.SyncSuccess, B.CardType, B.IncomeNumber, B.OutcomeNumber,
                            B.DocDate, B.RegDate, B.Summary, B.ToName, B.FromName, B.SignName, B.PerformName, B.Html);",
                card.OsedId, card.LinkDocId, /*0,*/ card.CardType, card.IncomeNumber, card.OutcomeNumber, card.DocDate,
                card.RegDate, card.Summary, card.ToName, card.FromName, card.SignName, card.PerformName, card.Html);

            CheckConnection();
            (new SqlCommand(cmd, sqlConnection)).ExecuteNonQuery();

            if (card.References != null)
            {
                foreach (var item in card.References)
                {
                    CheckConnection();
                    (new SqlCommand(string.Format(@"
                    MERGE [dbo].[4doc_OsedRefs] A
                        USING (VALUES ('{0}', {1}, '{2}', '{3}', '{4}')) B 
                        (OsedId, ResponseTo, RefNumber, Date, OrgName) ON (A.OsedId = B.OsedId)
                    WHEN MATCHED 
                        THEN UPDATE SET ResponseTo=B.ResponseTo, RefNumber=B.RefNumber, Date=B.Date, OrgName=B.OrgName
                    WHEN NOT MATCHED 
                        THEN INSERT (OsedId, ResponseTo, RefNumber, Date, OrgName)
                    VALUES (B.OsedId, B.ResponseTo, B.RefNumber, B.Date, B.OrgName);",
                    card.OsedId, item.Item1 ? 1 : 0, item.Item2, item.Item3, item.Item4),sqlConnection)).ExecuteNonQuery();
                }
            }

            Console.WriteLine("Saving all files...");
             bool syncSuccess = true;
            if (card.Files != null)
            {
                for (int i = 0; i < card.Files.Count; i++)
                {
                    bool fileSynhronized = false;
                    for (int repeatCount = 0; repeatCount < 40; repeatCount++)
                    {
                        byte[] strm;
                        if (File.Exists(Program.tempStoragePath + card.Files[i].Item2))
                        {
                            using (var stream = new FileStream(Program.tempStoragePath + card.Files[i].Item2, FileMode.Open, FileAccess.Read))
                            {
                                using (var reader = new BinaryReader(stream))
                                {
                                    strm = reader.ReadBytes((int)stream.Length);
                                    var sqlcmd = new SqlCommand(string.Format(@"
                                MERGE [dbo].[4doc_OsedFiles] A
                                    USING (VALUES ('{0}', '{1}', '{2}', @3)) B 
                                    (FileId, OsedId, Path, Data) ON (A.FileId = B.FileId)
                                WHEN MATCHED 
                                    THEN UPDATE SET OsedId = B.OsedId, Path = B.Path, Data = B.Data
                                WHEN NOT MATCHED 
                                    THEN INSERT (FileId, OsedId, Path, Data)
                                VALUES (B.FileId, B.OsedId, B.Path, B.Data);",
                                        card.Files[i].Item1, card.OsedId, card.Files[i].Item2), sqlConnection);
                                    sqlcmd.Parameters.Add("@3", SqlDbType.VarBinary, strm.Length).Value = strm;

                                    CheckConnection();
                                    sqlcmd.ExecuteNonQuery();

                                    fileSynhronized = true;
                                    Console.WriteLine("File {0} saved to base.", card.Files[i].Item2);
                                }
                            }
                            break;
                        }
                        else
                        {
                            System.Threading.Thread.Sleep(200);
                        }
                    }
                    if (File.Exists(Program.tempStoragePath + card.Files[i].Item2))
                    {
                        File.Delete(Program.tempStoragePath + card.Files[i].Item2);
                        Console.WriteLine("File {0} deleted from temp directory.", card.Files[i].Item2);
                    }
                    if (!fileSynhronized)
                    {
                        syncSuccess = false;
                        Console.WriteLine("NOT ALL FILES SYNCHONIZED IN CARD {0}", card.LinkDocId);
                    }
                }
            }

            CheckConnection();
            if (syncSuccess)
            {
                (new SqlCommand(string.Format(@"UPDATE [dbo].[4doc_Osed] SET
                    [SyncSuccess] = 1 WHERE [LinkDocId] = {0}", card.LinkDocId), sqlConnection)).ExecuteNonQuery();
            }
            //}
            //catch (Exception ex)
            //{
            //    //логирование
            //    Logger.WriteToBase(ex);
            //}
        }
        
        public void WriteLog(string message)
        {
            CheckConnection();
            (new SqlCommand(string.Format(@"INSERT INTO dbo.[4doc_OsedLogs]
                            (DateOfAction, Message) VALUES
                            (GETDATE(), '{0}')", message), sqlConnection)).ExecuteNonQuery();
        }

        public bool CardLinkExist(int linkId)
        {
            CheckConnection();
            return (new SqlCommand(string.Format(@"SELECT 1 FROM [dbo].[4doc_Osed] WHERE LinkDocId={0} AND SyncSuccess=1", linkId),
                sqlConnection)).ExecuteScalar() != null;
        }

        public Guid? FindCardGuid(int linkId)
        {
            CheckConnection();
            var rdr = (new SqlCommand(string.Format(@"SELECT OsedId FROM [dbo].[4doc_Osed] WHERE LinkDocId={0}", linkId),
                sqlConnection)).ExecuteReader();
            var guids = new List<Guid>();
            while (rdr.HasRows && rdr.Read())
            {
                guids.Add(rdr.GetGuid(0));
            }
            rdr.Close();
            if (guids.Count == 1)
                return guids[0];
            else
                return null;
        }

        public Guid? FindFileGuid(string path)
        {
            CheckConnection();
            var rdr = (new SqlCommand(string.Format(@"SELECT Path FROM [dbo].[4doc_OsedFiles] WHERE Path='{0}'", path),
                sqlConnection)).ExecuteReader();
            var guids = new List<Guid>();
            while (rdr.HasRows && rdr.Read())
            {
                guids.Add(rdr.GetGuid(0));
            }
            rdr.Close();
            if (guids.Count == 1)
                return guids[0];
            else
                return null;
        }

        public List<int> getNotSynchronized()
        {
            CheckConnection();
            var rdr = (new SqlCommand(@"SELECT LinkDocId FROM dbo.[4doc_Osed] WHERE SyncSuccess=0", sqlConnection)).ExecuteReader();
            var links = new List<int>();
            while (rdr.Read() && rdr.HasRows)
            {
                links.Add(rdr.GetInt32(0));
            }
            rdr.Close();
            return links;
        }

        /// <summary>
        /// Delete all osed bases
        /// </summary>
        public void DeleteAll()
        {
            try
            {
                CheckConnection();
                (new SqlCommand(@"DELETE FROM [dbo].[4doc_OsedFiles]", sqlConnection)).ExecuteNonQuery();
                (new SqlCommand(@"DELETE FROM [dbo].[4doc_OsedRefs]", sqlConnection)).ExecuteNonQuery();
                (new SqlCommand(@"DELETE FROM [dbo].[4doc_Osed]", sqlConnection)).ExecuteNonQuery();
            }
            catch (Exception ex)
            {
                //логирование
                Logger.WriteToBase(ex);
            }
        }

        /// <summary>
        /// Close connection
        /// </summary>
        public void Dispose()
        {
            sqlConnection.Close();
        }
    }
}
