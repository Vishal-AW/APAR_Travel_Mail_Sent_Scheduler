using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Security;
using System.Security.Cryptography;
using System.Text;
using System.Threading.Tasks;
using APAR_Travel_Mail_Sent_Scheduler.Models;
using UserInformation;
using MSC = Microsoft.SharePoint.Client;
namespace APAR_Travel_Mail_Sent_Scheduler.Models
{
    public static class CustomSharePointUtility
    {
        static UserOperation _UserOperation = new UserOperation();
        public static StreamWriter logFile;
        static byte[] bytes = ASCIIEncoding.ASCII.GetBytes("ZeroCool");
        public static string Decrypt(string cryptedString)
        {
            if (String.IsNullOrEmpty(cryptedString))
            {
                throw new ArgumentNullException("The string which needs to be decrypted can not be null.");
            }

            DESCryptoServiceProvider cryptoProvider = new DESCryptoServiceProvider();
            MemoryStream memoryStream = new MemoryStream(Convert.FromBase64String(cryptedString));
            CryptoStream cryptoStream = new CryptoStream(memoryStream, cryptoProvider.CreateDecryptor(bytes, bytes), CryptoStreamMode.Read);
            StreamReader reader = new StreamReader(cryptoStream);

            return reader.ReadToEnd();
        }
        public static MSC.ClientContext GetContext(string siteUrl)
        {
            try
            {
                AppConfiguration _AppConfiguration = GetSharepointCredentials(siteUrl);
                var securePassword = new SecureString();
                foreach (char c in _AppConfiguration.ServicePassword)
                {
                    securePassword.AppendChar(c);
                }
                var onlineCredentials = new MSC.SharePointOnlineCredentials(_AppConfiguration.ServiceUserName, securePassword);
                var context = new MSC.ClientContext(_AppConfiguration.ServiceSiteUrl);
                context.Credentials = onlineCredentials;
                return context;
            }
            catch (Exception ex)
            {
                WriteLog("Error in  CustomSharePointUtility.GetContext: "+ex.ToString());
                return null;
            }
        }
        public static void WriteLog(string logmsg)
        {
            // StreamWriter logFile;

            try
            {

            //    string LogString = DateTime.Now.ToString("dd/MM/yyyy HH:MM") + " " + logmsg.ToString();

                //  logFile.WriteLine(DateTime.Now);
                //  logFile.WriteLine(logmsg.ToString());
            //    logFile.WriteLine(LogString);

                //logFile.Close();
            }
            catch (Exception ex)
            {
             //   WriteLog(ex.ToString());

            }

        }

        public static AppConfiguration GetSharepointCredentials(string siteUrl)
        {
            AppConfiguration _AppConfiguration = new AppConfiguration();

            _AppConfiguration.ServiceSiteUrl = siteUrl;// _UserOperation.ReadValue("SP_Address");
            _AppConfiguration.ServiceUserName = _UserOperation.ReadValue("SP_USER_ID_Live");
            _AppConfiguration.ServicePassword = Decrypt(_UserOperation.ReadValue("SP_Password_Live"));

            return _AppConfiguration;
        }


        public static List<TravelVoucher> GetAll_TravelVoucherFromSharePoint(string siteUrl, string listName, string DaysDifference)
        {
            List<TravelVoucher> _retList = new List<TravelVoucher>();
            try
            {
                using (MSC.ClientContext context = CustomSharePointUtility.GetContext(siteUrl))
                {
                    if (context != null)
                    {
                        MSC.List list = context.Web.Lists.GetByTitle(listName);
                        MSC.ListItemCollectionPosition itemPosition = null;
                        while (true)
                        {   

                            var dataDateValue = DateTime.Now.AddDays(-Convert.ToInt32 (DaysDifference));
                            MSC.CamlQuery camlQuery = new MSC.CamlQuery();
                            camlQuery.ListItemCollectionPosition = itemPosition;
                            camlQuery.ViewXml = @"<View>
                                 <Query>
                                    <Where>
                                        <And>
                                            <Eq>
                                                <FieldRef Name='StatusCode'/>
                                                <Value Type='Text'>P</Value>
                                            </Eq> 
                                            <Leq><FieldRef Name='Modified'/><Value Type='DateTime'>" + dataDateValue.ToString("o") + "</Value></Leq>";                                                 
                                            camlQuery.ViewXml += @"</And>
                                    </Where>
                                </Query>
                                <RowLimit>4000</RowLimit>
                                <ViewFields>
                                <FieldRef Name='ID'/>
                                <FieldRef Name='ExpVoucherNo'/>
                                <FieldRef Name='CreatorName'/>
                                <FieldRef Name='CreatorDepartment'/>
                                <FieldRef Name='CreatorLocation'/>
                                <FieldRef Name='Destination'/>
                                <FieldRef Name='VisitDate'/>
                                <FieldRef Name='StatusCode'/> 
                                <FieldRef Name='StatusName'/>
                                <FieldRef Name='AssignUser'/>
                                <FieldRef Name='CreationDate'/>
                                <FieldRef Name='AssignDate'/>
                                <FieldRef Name='FunctionalHead'/>
                                <FieldRef Name='SequenceNo'/>
                                <FieldRef Name='EmployeeName'/>
                                <FieldRef Name='EmployeeNumber'/>
                                <FieldRef Name='Designation'/>
                                <FieldRef Name='ActionTaken'/>
                                <FieldRef Name='CompanyCode'/>
                                <FieldRef Name='DivisionName'/>
                                <FieldRef Name='Modified'/>
                                <FieldRef Name='TravelType'/>
                                </ViewFields></View>";
                            MSC.ListItemCollection Items = list.GetItems(camlQuery);

                            context.Load(Items);
                            context.ExecuteQuery();
                            itemPosition = Items.ListItemCollectionPosition;
                            foreach (MSC.ListItem item in Items)
                            {
                                _retList.Add(new TravelVoucher
                                {
                                    Id = Convert.ToInt32(item["ID"]),
                                    ExpVoucherNo = Convert.ToString(item["ExpVoucherNo"]).Trim(),
                                    CreatorName = Convert.ToString((item["CreatorName"] as Microsoft.SharePoint.Client.FieldUserValue).LookupValue),
                                    CreatorDepartment = Convert.ToString(item["CreatorDepartment"]).Trim(),
                                    CreatorLocation = Convert.ToString(item["CreatorLocation"]).Trim(),
                                    Destination = Convert.ToString(item["Destination"]).Trim(),
                                    VisitDate = Convert.ToString(item["VisitDate"]).Trim(),
                                    StatusCode = Convert.ToString(item["StatusCode"]).Trim(),
                                    StatusName = Convert.ToString(item["StatusName"]),
                                    AssignUser = item["AssignUser"] == null ? "" : Convert.ToString((item["AssignUser"] as Microsoft.SharePoint.Client.FieldUserValue[])[0].LookupId),
                                    FunctionalHead = item["FunctionalHead"] == null ? "" : Convert.ToString(item["FunctionalHead"]).Trim(),
                                    CreationDate = Convert.ToString(item["CreationDate"]).Trim(),
                                    AssignDate = Convert.ToString(item["AssignDate"]).Trim(),
                                    SequenceNo = Convert.ToString(item["SequenceNo"]).Trim(),
                                    EmployeeName = Convert.ToString(item["EmployeeName"]).Trim(),
                                    EmployeeNumber = Convert.ToString(item["EmployeeNumber"]).Trim(),
                                    Designation = Convert.ToString(item["Designation"]).Trim(),
                                    ActionTaken = Convert.ToString(item["ActionTaken"]).Trim(),
                                    CompanyCode = Convert.ToString(item["CompanyCode"]),
                                    DivisionName = Convert.ToString(item["DivisionName"]),
                                    Modified = Convert.ToString(item["Modified"]),
                                    TravelType = Convert.ToString(item["TravelType"]),
                                });
                            }
                            if (itemPosition == null)
                            {
                                break; // TODO: might not be correct. Was : Exit While
                            }

                        }
                    }
                }
            }
            catch (Exception ex)
            {
                CustomSharePointUtility.WriteLog("Error in  GetAll_EmployeeDetailsFromSharePoint()" + " Error:" + ex.Message);
            }
            return _retList;
        }
        //public static void sample()
        //{
        //    List<TravelVoucher> data = new List<TravelVoucher>();
        //  List<Mailing> mail=EmailData(data, "", "");
        //}

        public static bool EmailData(List<TravelVoucher> updationList, string siteUrl, string listName)
        {
            bool retValue = false;
            try
            {

                using (MSC.ClientContext context = CustomSharePointUtility.GetContext(siteUrl))
                {
                    //List<Mailing> varx = new List<Mailing>();

                    MSC.List list = context.Web.Lists.GetByTitle(listName);
                    for (var i = 0; i < updationList.Count; i++)
                    {
                        var updateList = updationList.Skip(i).Take(1).ToList();
                        if (updateList != null && updateList.Count > 0)
                        {
                            foreach (var updateItem in updateList)
                            {
                                MSC.ListItem listItem = null;
                             
                                    MSC.ListItemCreationInformation itemCreateInfo = new MSC.ListItemCreationInformation();
                                    listItem = list.AddItem(itemCreateInfo);
                                
                                var obj = new Object();
                                //Mailing data = new Mailing();
                                
                                //var _From = "";
                                var _To = "";
                                //var _Cc = "";
                                var _Body = "";
                                var _Subject = "";
                                if (updateItem.SequenceNo == "1")
                                {
                                    _To = updateItem.FunctionalHead;
                                }
                                else
                                {
                                    _To = updateItem.AssignUser;
                                }
                                _Subject = "Gentle Reminder"; // + updateItem.ExpVoucherNo + " Travel Voucher Approval is Pending
                                _Body += "Dear User, <br><br>This is to inform you that below request is pending for your Approval.";
                                _Body += "<br><b>Workflow Name :</b> Travel Voucher ";
                                _Body += "<br><b>Voucher No :</b>  " + updateItem.ExpVoucherNo;
                                _Body += "<br><b>Date of Creation :</b>  " + updateItem.CreationDate;
                                _Body += "<br><b>Employee : </b> " + updateItem.CreatorName;
                                _Body += "<br><b>Designation :</b> " + updateItem.Designation;

                                _Body += "<br><b>Department :</b> " + updateItem.CreatorDepartment;
                                _Body += "<br><b>Location :</b> " + updateItem.CreatorLocation;
                                if (updateItem.SequenceNo == "1")
                                {
                                    _Body += "<br><b>Status :</b> Pending With Functional Head";
                                }
                                else if (updateItem.SequenceNo == "5")
                                {
                                    _Body += "<br><b>Status :</b> Pending With Travel Desk";
                                }
                                else if (updateItem.SequenceNo == "2")
                                {
                                    _Body += "<br><b>Status :</b> Pending With Internal Audit";
                                }
                                else if (updateItem.SequenceNo == "3")
                                {
                                    _Body += "<br><b>Status :</b> Pending With Accounts";
                                }
                                _Body += "<br><h3>Kindly provide your approval</h3>";
                                _Body += "<br><h3>For Approval Please Click in the below link</h3>";
                                if (updateItem.SequenceNo == "1") {
                                    _Body += "<br><a href=\"https://aparindltd.sharepoint.com/TravelVoucher/SitePages/PendingFunctionalHead.aspx\">View Link</a>";
                                }
                                else if (updateItem.SequenceNo == "5") {
                                    _Body += "<br><a href=\"https://aparindltd.sharepoint.com/TravelVoucher/SitePages/PendingTravelDesk.aspx\">View Link</a>";
                                }
                                else if (updateItem.SequenceNo == "2")
                                {
                                    _Body += "<br><a href=\"https://aparindltd.sharepoint.com/TravelVoucher/SitePages/Pending%20Request.aspx\">View Link</a>";
                                }
                                else if (updateItem.SequenceNo == "3")
                                {
                                    _Body += "<br><a href=\"https://aparindltd.sharepoint.com/TravelVoucher/SitePages/PendingRequestAccounts.aspx\">View Link</a>";
                                }

                                //data.MailTo = _From;
                                //data.MailTo = _To;
                                //data.MailCC = _Cc;
                                //data.MailSubject = _Subject;
                                //data.MailBody = _Body;
                                //varx.Add(data);
                                listItem["ToUser"] = _To;
                                listItem["MailSubject"] = _Subject;
                                listItem["MailBody"] = _Body;
                                listItem.Update();
                            }
                            try
                            {
                                context.ExecuteQuery();
                                retValue = true;

                            }
                            catch (Exception ex)
                            {
                                CustomSharePointUtility.WriteLog(string.Format("Error in  InsertUpdate_EmployeeMaster ( context.ExecuteQuery();): Error ({0}) ", ex.Message));
                                return false;
                                //continue;
                            }
                        }
                    }
                }

            }
            catch (Exception ex)
            {
                CustomSharePointUtility.WriteLog(string.Format("Error in  InsertUpdate_EmployeeMaster: Error ({0}) ", ex.Message));
            }
            return retValue;

        }
    }
    public class AppConfiguration
    {
        public string ServiceSiteUrl;
        public string ServiceUserName;
        public string ServicePassword;
    }
}
