using System;
using System.Collections.Generic;
using System.Linq;
using Dapper;
using Microsoft.Azure.WebJobs;
using Microsoft.Data.SqlClient;
using Microsoft.Extensions.Logging;
using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.News.DataModel;
using SharepointBirthdayUtility;
using SharePointBirthDayUtility.Models;


namespace SharePointBirthDayUtility
{
    public class Program
    {
        private static string SPSiteUrl { get { return Environment.GetEnvironmentVariable("SPSiteUrl"); } }
        private static string SPClientId { get { return Environment.GetEnvironmentVariable("SPClientId"); } }
        private static string SPClientSecret { get { return Environment.GetEnvironmentVariable("SPClientSecret"); } }
        private static string SPFolderName { get { return Environment.GetEnvironmentVariable("SPFolderName"); } }
        private static string ConString { get { return Environment.GetEnvironmentVariable("SqlConnectionString"); } }

        public static int UpdatedUserCount = 0;
        public static List<EmployeeDetailModel> emailDto = new();

        [FunctionName("Program")]
        public static void Run([TimerTrigger("0 */1 * * * *")] TimerInfo myTimer, ILogger log)
        {
            try
            {
                log.LogDebug(Environment.NewLine + Environment.NewLine + "Trace Log:" + Environment.NewLine + Environment.NewLine + "--Log Entry : " + DateTime.Now.ToString() + " " + Environment.NewLine);

                //string dateString = DateTime.Now.ToString("yyyy-MM-dd");
                int intResult = GetEmployeeDetailListFromDB(out List<EmployeeDetailModel> empLST, out Exception exObj);
                if (intResult == -1)
                {
                    log.LogDebug($"Database issue " + exObj.Message);
                    return;
                }

                //-------------------
                log.LogDebug($"User Count: {empLST.Count} ");
                log.LogDebug((empLST.Count == 0) ? "--No Employee record found--" : $"{empLST.Count}" + "" + Environment.NewLine);
                //-------------------

                intResult = ConnectSP(out ClientContext clientContext, out exObj);
                if (intResult == -1)
                {
                    log.LogDebug($"SP issue " + exObj.Message);
                    return;
                }
                //-------------------

                if (!clientContext.Web.ListExists(SPFolderName))
                {
                    log.LogDebug($"{SPFolderName} SP list not exists");
                    return;
                }

                //-------------------
                int count = 1;
                List oList = clientContext.Web.Lists.GetByTitle(SPFolderName);
                foreach (var q2employees in empLST)
                {
                    int intOldItemID = IsUserExist(ref clientContext, ref oList, q2employees, ref log);
                    if (intOldItemID > 0)
                    {
                        UpdateListItem(ref clientContext, ref oList, intOldItemID, q2employees, ref log);
                        count += 1;
                    }
                    else
                    {
                        UpdatedUserCount += 1;
                        log.LogDebug($"Employee Not Found : {UpdatedUserCount}. {q2employees.EmpCode}" + $" {q2employees.EmailAddress}" + Environment.NewLine);
                    }
                }
                EmailHelper.SendApiErrorEmail(emailDto, ref log);
                log.LogDebug($"Total Employee Updated : {count}" + Environment.NewLine);
            }

            catch (Exception ex)
            {
                log.LogDebug("--Error : --" + ex.Message + Environment.NewLine);
            }
        }

        private static int ConnectSP(out ClientContext clientContext, out Exception exObj)
        {
            clientContext = null;
            exObj = new Exception();
            try
            {
                using ClientContext context = new PnP.Framework.AuthenticationManager().GetACSAppOnlyContext(SPSiteUrl, SPClientId, SPClientSecret);
                context.Load(context.Web.Folders);
                context.ExecuteQueryAsync().ConfigureAwait(false).GetAwaiter().GetResult();

                clientContext = context;
            }
            catch (Exception ex)
            {
                exObj = ex;
                return -1;
            }
            return 1;
        }

        private static int GetEmployeeDetailListFromDB(out List<EmployeeDetailModel> empLST, out Exception exObj)
        {
            empLST = new List<EmployeeDetailModel>();
            exObj = new Exception();
            try
            {
                //Database Queries   
                using var con = new SqlConnection(ConString);
                string commandText = "{{SQL QUERY}}";
                empLST = con.Query<EmployeeDetailModel>(commandText).OrderBy(x => x.FirstName).Take(100).ToList();
                con.Close();
            }
            catch (Exception ex)
            {
                exObj = ex;
                return -1;
            }
            return 1;
        }

        private static void UpdateListItem(ref ClientContext clientContext, ref List oList, Int32 ItemID, EmployeeDetailModel _q2employees, ref ILogger log)
        {
            UpdatedUserCount += 1;
            try
            {
                ListItem oListItem = oList.GetItemById(ItemID);
                oListItem["Anniversary"] = Convert.ToDateTime(_q2employees.Datehire);
                oListItem["Birthday"] = Convert.ToDateTime(_q2employees.DOB);
                oListItem["EmployeeID"] = _q2employees.EmpCode;
                log.LogDebug($"Employee Updated : {UpdatedUserCount}. {_q2employees.EmpCode}" + $" {_q2employees.EmailAddress}" + Environment.NewLine);
                oListItem.Update();
                clientContext.ExecuteQuery();
            }
            catch (Exception ex)
            {
                log.LogDebug($"--Error : {UpdatedUserCount}" + ex.Message + Environment.NewLine);
            }
        }

        private static int IsUserExist(ref ClientContext clientContext, ref List oList, EmployeeDetailModel empInfo, ref ILogger log)
        {
            try
            {
                int checkRecord = RecordExists(ref clientContext, ref oList, RecordType.WorkEmail, empInfo.EmailAddress);
                if (checkRecord > 0) { return checkRecord; }

                checkRecord = RecordExists(ref clientContext, ref oList, RecordType.EmployeeID, empInfo.EmpCode);
                if (checkRecord > 0) { return checkRecord; }
                //---------------
                AddEmplyoeeDetail(empInfo);
            }
            catch (Exception ex)
            {
                AddEmplyoeeDetail(empInfo);
                log.LogDebug("--Error : --" + ex.Message + Environment.NewLine);
            }
            return 0;
        }

        private enum RecordType
        {
            WorkEmail,
            EmployeeID
        }

        private static int RecordExists(ref ClientContext clientContext, ref List oList, RecordType recordType, string value)
        {
            int ItemID = 0;
            var camlQuery = new CamlQuery();
            if (recordType == RecordType.WorkEmail)
            {
                camlQuery.ViewXml = "<View><Query><Where><Eq><FieldRef Name='WorkEmail'/><Value Type='Text'>" + value + "</Value></Eq></Where></Query></View>";
            }
            else if (recordType == RecordType.EmployeeID)
            {

                camlQuery.ViewXml = "<View><Query><Where><Eq><FieldRef Name='EmployeeID'/><Value Type='Text'>" + value + "</Value></Eq></Where></Query></View>";
            }
            ListItemCollection collListItem = oList.GetItems(camlQuery);
            clientContext.Load(collListItem);
            clientContext.ExecuteQuery();
            if (collListItem.Count > 0)
            {
                foreach (ListItem item in collListItem) { ItemID = item.Id; }
            }
            return ItemID;

        }

        private static void AddEmplyoeeDetail(EmployeeDetailModel empInfo)
        {
            emailDto.Add(new EmployeeDetailModel
            {
                EmailAddress = empInfo.EmailAddress,
                EmpCode = empInfo.EmpCode,
                FirstName = empInfo.FirstName,
                LastName = empInfo.LastName
            });
        }
    }
}

