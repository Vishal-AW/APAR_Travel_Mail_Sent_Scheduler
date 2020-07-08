using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using APAR_Travel_Mail_Sent_Scheduler.Models;
using System.Configuration;

namespace APAR_Travel_Mail_Sent_Scheduler
{
    class Program
    {
        static void Main(string[] args)
        {
            //Console.WriteLine("*********************************************");
            //Console.WriteLine("Reminder Mail starts : " + DateTime.Now.ToString());
            //Console.WriteLine("*********************************************");
            List<TravelVoucher> SPTravelVoucher = null;
            try
            {
                var siteUrl = ConfigurationManager.AppSettings["SP_Address_Live"];
                string TestingTravelHeaderList = ConfigurationManager.AppSettings["TestingTravelHeaderList"];
                string EmailList = ConfigurationManager.AppSettings["EmailList"];
                string DaysDifference = ConfigurationManager.AppSettings["DaysDifference"];
                //string query = SQLUtility.ReadQuery("EmployeeMasterQuery.txt");
                SPTravelVoucher = new List<TravelVoucher>();
                //Task task_SPEmployeeMaster = Task.Run(() => SPTravelVoucher = CustomSharePointUtility.GetAll_TravelVoucherFromSharePoint(siteUrl, TestingTravelHeaderList));
                SPTravelVoucher = CustomSharePointUtility.GetAll_TravelVoucherFromSharePoint(siteUrl, TestingTravelHeaderList, DaysDifference);
                //List<TravelVoucher> empMasterFinal = new List<TravelVoucher>();
                List<TravelVoucher> empMasterFinal = SPTravelVoucher;
                if (empMasterFinal.Count > 0)
                {
                    //CustomSharePointUtility.WriteLog("Voucher data successfully.");
                    //Console.WriteLine("Employee data synchronized successfully.");
                    var success = CustomSharePointUtility.EmailData(empMasterFinal, siteUrl, EmailList);
                    if (success)
                    {
                        //CustomSharePointUtility.WriteLog("Reminder Mail Sent Successfully.");
                        //Console.WriteLine("Reminder Mail Sent Successfully.");
                    }
                }
                else
                {
                    //CustomSharePointUtility.WriteLog("No Pending Records.");
                    //Console.WriteLine("No Pending Records.");
                }
            }
            catch (Exception ex)
            {
                //CustomSharePointUtility.WriteLog("Error in scheduler : " + ex.StackTrace);
                //Console.WriteLine("Error in scheduler : " + ex.StackTrace);
            }
            finally
            {
                //CustomSharePointUtility.WriteLog("*********************************************");
                //CustomSharePointUtility.WriteLog("Reminder Mail ends : " + DateTime.Now.ToString());
                //CustomSharePointUtility.WriteLog("*********************************************");
            //    Console.WriteLine("*********************************************");
             //   Console.WriteLine("Reminder Mail ends : " + DateTime.Now.ToString());
            //    Console.WriteLine("*********************************************");
              //  CustomSharePointUtility.logFile.Close();
                //Console.ReadKey();

            }
        }
    }
}
