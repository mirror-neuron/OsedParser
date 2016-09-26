using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OsedParser
{
    public static class Logger
    {
        public static SQL sql;
        public static void WriteToBase(Exception ex)
        {
            sql.WriteLog(ex.StackTrace);
        }
    }
}