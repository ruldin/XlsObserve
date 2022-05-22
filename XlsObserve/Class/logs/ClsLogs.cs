using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace XlsObserve.Class.logs
{
    public static class ClsLogs
    {
        public static readonly int info = 1;
        public static readonly int actions = 2;

        public static void StatusTrace(string msg,int level)
        {
            Console.WriteLine(msg);
        }


        public  static void ErrorLog(string message)
        {
            Console.WriteLine($"error {message}");
        }

        




    }
}
