using System;
using System.Configuration;

namespace Epplus_Export
{
    class Program
    {
        private static IKDE_Email emailer = new IKDE_Email();
        public static AppSettingsReader apReader = new AppSettingsReader();
        private static DAL dal = new DAL();

        static void Main(string[] args)
        {
            Console.WriteLine("\nIBS_KWS_DailyTransaction_Extract...");
            dal.RetriveTransactions();   
        }
    }
}
