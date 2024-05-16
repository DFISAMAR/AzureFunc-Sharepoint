using Microsoft.Extensions.Logging;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using SharePointBirthDayUtility.Models;
using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Net.Mail;
using System.Net.Mime;

namespace SharepointBirthdayUtility
{
    public static class EmailHelper
    {
        public static string EmailContent => @"Hello Team,<br/><br/> Birthday Anniversay Utility run successfully.<br/>";
        public static string EmailContentError => @"Hello Team,<br/><br/> Birthday Anniversay Utility run with some failure.<br/>";
        private static string SmtpUserName { get { return Environment.GetEnvironmentVariable("smtpUserName"); } }
        private static string SmtpPasswrod { get { return Environment.GetEnvironmentVariable("smtpPasswrod"); } }
        private static string SmtpServer { get { return Environment.GetEnvironmentVariable("smtpServer"); } }
        private static string SmtpPort { get { return Environment.GetEnvironmentVariable("smtpPort"); } }
        private static string SendTo { get { return Environment.GetEnvironmentVariable("sendTo"); } }

        public static void SendApiErrorEmail(List<EmployeeDetailModel> data, ref ILogger log)
        {
            try
            {
                IWorkbook workbook = new XSSFWorkbook();
                ISheet excelSheet = workbook.CreateSheet("sheet1");
                IRow row = excelSheet.CreateRow(0);
                row.CreateCell(0).SetCellValue("First Name");
                excelSheet.SetColumnWidth(0, 2500);

                row.CreateCell(1).SetCellValue("Last Name");
                excelSheet.SetColumnWidth(1, 6000);

                row.CreateCell(2).SetCellValue("Email Address");
                excelSheet.SetColumnWidth(2, 6000);

                row.CreateCell(3).SetCellValue("Emp Code");
                excelSheet.SetColumnWidth(3, 6000);

                var rowIndex = 1;
                foreach (var item in data)
                {
                    row = excelSheet.CreateRow(rowIndex);
                    row.CreateCell(0).SetCellValue(item.FirstName);
                    row.CreateCell(1).SetCellValue(item.LastName);
                    row.CreateCell(2).SetCellValue(item.EmailAddress);
                    row.CreateCell(3).SetCellValue(item.EmpCode);
                    rowIndex++;
                }

                var byteArray = new List<byte>();
                using var stream = new MemoryStream();
                workbook.Write(stream, false);
                byteArray.AddRange(stream.ToArray());
                var doc = stream.ToArray();
                var checkMailStatus = SendMail(data, byteArray.ToArray(), out Exception exObj);
                if (checkMailStatus != 1)
                {
                    log.LogDebug(exObj.Message);
                    return;
                }
            }
            catch (Exception ex)
            {
                log.LogDebug("Error: " + ex.Message);
            }
        }

        private static int SendMail(List<EmployeeDetailModel> data, byte[] byteArray, out Exception exObj)
        {
            exObj = new Exception();
            try
            {
                using var mailMessage = new MailMessage();
                string strUserName = SmtpUserName;
                string strPassword = SmtpPasswrod;
                string strSmtpServer = SmtpServer;
                var smtpPort = Convert.ToInt16(SmtpPort);
                mailMessage.From = new MailAddress(strUserName);
                string strTo = SendTo;
                foreach (string strEmail in strTo.Split(new[] { ";" }, StringSplitOptions.RemoveEmptyEntries)) { mailMessage.To.Add(strEmail); }

                mailMessage.Priority = MailPriority.High;
                mailMessage.Subject = "Birthday-Anniversary Sync";

                string strEmailBody = data.Count > 0 ? string.Format(EmailContentError, "") : string.Format(EmailContent, "");
                mailMessage.Body = (strEmailBody != null && strEmailBody != "") ? strEmailBody : "";

                if (data.Count > 0)
                {
                    mailMessage.Attachments.Add(new Attachment(new MemoryStream(byteArray), "Employee Failed Data.xlsx", MediaTypeNames.Application.Octet));
                }

                mailMessage.IsBodyHtml = true;
                var smtpClient = new SmtpClient(strSmtpServer, smtpPort) { EnableSsl = true, Credentials = new System.Net.NetworkCredential(strUserName, strPassword) };

                smtpClient.Send(mailMessage);
                return 1;
            }
            catch (Exception ex)
            {
                exObj = ex;
                return -1;
            }
        }
    }
}
