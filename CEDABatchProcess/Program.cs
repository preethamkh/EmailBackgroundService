using System;
using System.Collections.Generic;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using System.Data.SqlClient;
using System.Configuration;
using System.IO;
using System.Net.Mail;
using System.Net;
using CEDABatchProcess.RSExecution2005;
using System.Security.Principal;
using System.Web.Services;
using System.Web.Services.Protocols;

namespace CEDABatchProcess
{
    class Program
    {
        static string coordName = "", coordPhone = "", coordEmail = "", eventName = "", contactName = "";
        static void Main(string[] args)
        {
            //SendEventInvoiceEmail();
            SendSSRSPDFEmail();
        }

        private static void SendSSRSPDFEmail()
        {
            string recordsToUpdate = "";
            String connString = ConfigurationManager.ConnectionStrings["LocalDB"].ToString();

            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(connString))
                {
                    // keep the same connection open for multiple sql commands, reducing DB hits
                    sqlConnection.Open(); 

                    // hold all registered events to check if the invoice is to be sent or not
                    List<String> allRegisteredEvents = new List<String>();

                    using (SqlCommand cmd = new SqlCommand("select top 5 * from MailEventInvoice where sentflag = @flag", sqlConnection))
                    {
                        cmd.Parameters.AddWithValue("flag", false);
                        string contactId = "", eventCode = "", coordinatorId = "", emailTo = "", record = "", productCode = "";

                        using (SqlDataReader reader = cmd.ExecuteReader(System.Data.CommandBehavior.Default))
                        {
                            while (reader.Read())
                            {
                                bool calendarOnly = false;
                                contactId = reader["BT_ID"].ToString();
                                eventCode = reader["EventCode"].ToString();
                                coordinatorId = reader["EventCoordinatorId"].ToString();
                                emailTo = reader["EmailToList"].ToString();
                                record = reader["ID"].ToString();
                                productCode = reader["ProductCode"].ToString();

                                // Get all registered events for this BT_ID
                                using (SqlConnection conn = new SqlConnection(ConfigurationManager.ConnectionStrings["DataSource.iMIS.Connection"].ToString()))
                                {
                                    conn.Open();

                                    using (SqlCommand storedProc = new SqlCommand("evo.sp_GetAllRegisteredEvents", conn))
                                    {
                                        storedProc.CommandType = CommandType.StoredProcedure;
                                        storedProc.Parameters.AddWithValue("@ContactId", contactId);

                                        using (SqlDataReader dr = storedProc.ExecuteReader())
                                        {
                                            while (dr.Read())
                                            {
                                                if(Decimal.Parse(dr["TotalCharges"].ToString()) > 0)
                                                    allRegisteredEvents.Add(dr["EventCode"].ToString().ToUpper());
                                            }
                                        }
                                    }
                                }

                                // Check if the event is present in the list of all registered events, if not just send the Calendar appointment via email
                                if(!allRegisteredEvents.Contains(eventCode))
                                    calendarOnly = true;         

                                ReportExecutionService rs = new ReportExecutionService();
                                rs.Credentials = new NetworkCredential(ConfigurationManager.AppSettings["SSRS.Username"].ToString(), ConfigurationManager.AppSettings["SSRS.Password"].ToString(), "CEDA");
                                rs.Url = ConfigurationManager.AppSettings["SSRS.WebService"].ToString();

                                // Render arguments
                                byte[] result = null;
                                string reportPath = ConfigurationManager.AppSettings["SSRS.ReportPath"].ToString();
                                string format = "PDF";
                                string historyID = null;
                                string devInfo = @"<DeviceInfo><Toolbar>False</Toolbar></DeviceInfo>";

                                // Prepare report parameter.
                                RSExecution2005.ParameterValue[] parameters = new RSExecution2005.ParameterValue[2];
                                parameters[0] = new RSExecution2005.ParameterValue();
                                parameters[0].Name = "ContactId";
                                parameters[0].Value = contactId;
                                parameters[1] = new RSExecution2005.ParameterValue();
                                parameters[1].Name = "meeting";
                                parameters[1].Value = eventCode;

                                string encoding;
                                string mimeType;
                                string extension;
                                RSExecution2005.Warning[] warnings = null;
                                string[] streamIDs = null;

                                ExecutionInfo execInfo = new ExecutionInfo();
                                ExecutionHeader execHeader = new ExecutionHeader();

                                rs.ExecutionHeaderValue = execHeader;

                                execInfo = rs.LoadReport(reportPath, historyID);

                                rs.SetExecutionParameters(parameters, "en-us");
                                String SessionId = rs.ExecutionHeaderValue.ExecutionID;

                                //Console.WriteLine("SessionID: {0}", rs.ExecutionHeaderValue.ExecutionID);

                                try
                                {
                                    if (!calendarOnly)
                                    {
                                        result = rs.Render(format, devInfo, out extension, out mimeType, out encoding, out warnings, out streamIDs);
                                        execInfo = rs.GetExecutionInfo();
                                    }
                                    else
                                        result = null;                                    
                                }
                                catch (SoapException e)
                                {
                                    Console.WriteLine(e.Detail.OuterXml);
                                }

                                try
                                {
                                    using (SmtpClient client = new SmtpClient())
                                    {
                                        Int32 iPort = 0;
                                        Int32.TryParse(ConfigurationManager.AppSettings["Email.Port"].ToString(), out iPort);
                                        if (iPort > 0)
                                            client.Port = iPort;

                                        client.Host = ConfigurationManager.AppSettings["Email.Host"].ToString();
                                        client.DeliveryMethod = SmtpDeliveryMethod.Network;
                                        client.EnableSsl = Boolean.Parse(ConfigurationManager.AppSettings["Email.EnableSSL"].ToString());
                                        //client.Port = Int32.Parse(ConfigurationManager.AppSettings["Email.Port"].ToString());

                                        String username = ConfigurationManager.AppSettings["Email.Username"].ToString();
                                        String password = ConfigurationManager.AppSettings["Email.Password"].ToString();

                                        if (!string.IsNullOrEmpty(username) && !string.IsNullOrEmpty(password))
                                        {
                                            client.UseDefaultCredentials = false;
                                            client.Credentials = new NetworkCredential(username, password);
                                        }
                                        else
                                            client.UseDefaultCredentials = true;


                                        if (!string.IsNullOrEmpty(emailTo))
                                        {
                                            using (MailMessage message = new MailMessage())
                                            {
                                                StringBuilder emailBody = new StringBuilder();
                                                message.From = new MailAddress("noreply@ceda.com.au");
                                                var bccMails = ConfigurationManager.AppSettings["BCCList"].ToString().Split(';');
                                                string replyState = "";

                                                foreach (var email in bccMails)
                                                {
                                                    message.Bcc.Add(email.Trim());
                                                }

                                                //POCS-25
                                                if (eventCode.Trim().ToLower().Contains("nec"))
                                                {
                                                    message.ReplyToList.Add(new MailAddress(ConfigurationManager.AppSettings["National.From"].ToString()));
                                                    message.Bcc.Add(new MailAddress(ConfigurationManager.AppSettings["National.From"].ToString()));

                                                    replyState = "false";
                                                }
                                                else  if(eventCode.Trim().ToLower().StartsWith("nt"))
                                                {
                                                    message.ReplyToList.Add(new MailAddress(ConfigurationManager.AppSettings["SA.From"].ToString()));
                                                    message.Bcc.Add(new MailAddress(ConfigurationManager.AppSettings["SA.From"].ToString()));

                                                    replyState = "false";
                                                }

                                                
                                                if (!string.IsNullOrEmpty(eventCode.ToLower()) && !replyState.Equals("false"))
                                                {
                                                    replyState = eventCode[0].ToString().ToLower();

                                                    switch (replyState)
                                                    {
                                                        case "q":
                                                            message.ReplyToList.Add(new MailAddress(ConfigurationManager.AppSettings["QLD.From"].ToString()));
                                                            message.Bcc.Add(new MailAddress(ConfigurationManager.AppSettings["QLD.From"].ToString()));
                                                            break;

                                                        case "n":
                                                            message.ReplyToList.Add(new MailAddress(ConfigurationManager.AppSettings["NSW.From"].ToString()));
                                                            message.Bcc.Add(new MailAddress(ConfigurationManager.AppSettings["NSW.From"].ToString()));
                                                            break;

                                                        case "a":
                                                            message.ReplyToList.Add(new MailAddress(ConfigurationManager.AppSettings["NSW.From"].ToString()));
                                                            message.Bcc.Add(new MailAddress(ConfigurationManager.AppSettings["NSW.From"].ToString()));
                                                            break;

                                                        case "v":
                                                            message.ReplyToList.Add(new MailAddress(ConfigurationManager.AppSettings["VIC.From"].ToString()));
                                                            message.Bcc.Add(new MailAddress(ConfigurationManager.AppSettings["VIC.From"].ToString()));
                                                            break;

                                                        case "t":
                                                            message.ReplyToList.Add(new MailAddress(ConfigurationManager.AppSettings["VIC.From"].ToString()));
                                                            message.Bcc.Add(new MailAddress(ConfigurationManager.AppSettings["VIC.From"].ToString()));
                                                            break;

                                                        case "s":
                                                            message.ReplyToList.Add(new MailAddress(ConfigurationManager.AppSettings["SA.From"].ToString()));
                                                            message.Bcc.Add(new MailAddress(ConfigurationManager.AppSettings["SA.From"].ToString()));
                                                            break;

                                                        case "w":
                                                            message.ReplyToList.Add(new MailAddress(ConfigurationManager.AppSettings["WA.From"].ToString()));
                                                            message.Bcc.Add(new MailAddress(ConfigurationManager.AppSettings["WA.From"].ToString()));
                                                            break;

                                                        case "cop":
                                                            message.ReplyToList.Add(new MailAddress(ConfigurationManager.AppSettings["Copland.From"].ToString()));
                                                            message.Bcc.Add(new MailAddress(ConfigurationManager.AppSettings["Copland.From"].ToString()));
                                                            break;

                                                        default:
                                                            message.ReplyToList.Add(new MailAddress("noreply@ceda.com.au"));
                                                            break;
                                                    }
                                                }                                                

                                                message.To.Add(emailTo);
                                                
                                                message.Subject = reader["InvoiceSubject"].ToString();
                                                message.IsBodyHtml = true;
                                                message.BodyEncoding = Encoding.GetEncoding(1254);
                                                
                                                message.Attachments.Add(OutlookCalendarAppointment(eventCode, coordinatorId, contactId, productCode));
                                                emailBody.Append("<html><body><div style='font-size: 16px; font-family: Calibri, Helvetica, Arial, Sans-Serif;'>Dear " + contactName + ",<br /><br />");

                                                if (result != null)
                                                {
                                                    message.Attachments.Add(new Attachment(new MemoryStream(result), "Invoice.pdf"));
                                                    emailBody.Append("Please find attached the invoice and calendar appointment for your recent registration for the CEDA event:<br /><br />");
                                                }
                                                else
                                                    emailBody.Append("Please find attached the calendar appointment for your recent registration for the CEDA event:<br /><br />");

                                                emailBody.Append("<span style='font-weight: bold; font-size:18px;'>" + eventName.ToUpper() + "<span><br /><br />");
                                                emailBody.Append("<span style='font-weight: normal; font-size:16px;'>For any questions, please contact the event co-ordinator:<span><br /><br />");
                                                emailBody.Append("<span style='font-weight: bold;'>" + coordName + "</span><br />");
                                                emailBody.Append("<span style='font-weight: bold;'>Phone: </span><span 'font-weight: normal;'>" + coordPhone + "</span><br />");
                                                emailBody.Append("<span style='font-weight: bold;'>Email: </span><span 'font-weight: normal;'>" + coordEmail + "</span><br />");
                                                emailBody.Append("<br /><br /><span 'font-weight: normal;'>[THIS IS AN AUTOMATED MESSAGE - PLEASE DO NOT REPLY DIRECTLY TO THIS EMAIL]</span>");
                                                emailBody.Append("</div></body></html>");
                                                message.Body = emailBody.ToString();

                                                client.Send(message);

                                                // UPDATE RECORD IN DB IF EMAIL IS SENT SUCCESSFULLY
                                                if (string.IsNullOrEmpty(recordsToUpdate))
                                                {
                                                    recordsToUpdate = record;
                                                }
                                                else
                                                    recordsToUpdate += "," + record;
                                            }
                                        }
                                    }
                                }
                                catch (Exception e)
                                {
                                    Console.WriteLine(e.Message);
                                }
                            }
                        }

                        // UPDATE DB RECORDS AFTER E-MAILS ARE SENT
                        if (!string.IsNullOrEmpty(recordsToUpdate))
                        {
                            string[] recordIds = recordsToUpdate.Split(',');
                            foreach (var variable in recordIds)
                            {
                                using (SqlCommand updateCommand = new SqlCommand("update MailEventInvoice set sentflag = @sentflag, senttimestamp = @senttimestamp where id = @id", sqlConnection))
                                {
                                    updateCommand.Parameters.AddWithValue("id", Int32.Parse(variable));
                                    updateCommand.Parameters.AddWithValue("senttimestamp", DateTime.Now);
                                    updateCommand.Parameters.AddWithValue("sentflag", true);
                                    var result = updateCommand.ExecuteNonQuery();
                                }
                            }
                        }
                    }
                }
                Environment.Exit(0);
            }
            catch (Exception)
            {
                Environment.Exit(1);
            }
        }


        private static Attachment OutlookCalendarAppointment(string eventCode, string coordinatorId, string contactId, string productCode)
        {
            Attachment objAttachment = null;

            try
            {
                // Check timezone for outlook calendar settings

                string eventTZ = "AUS Eastern Standard Time";
                string eventCodeState = eventCode.ToLower();
                eventCodeState = eventCodeState.Substring(0, 1);

                switch (eventCodeState)
                {
                    case "s": // SA
                        eventTZ = ConfigurationManager.AppSettings["SA"].ToString();
                        break;
                    case "w": // WA
                        eventTZ = ConfigurationManager.AppSettings["WA"].ToString();
                        break;
                    case "q": // QLD
                        eventTZ = ConfigurationManager.AppSettings["QLD"].ToString();
                        break;
                    default:
                        eventTZ = ConfigurationManager.AppSettings["DefaultTimeZone"].ToString();
                        break;
                }


                // Local time zone info. to be used to enable correct calendar entry for people registering from any TZ (in / outside AUS)
                TimeZoneInfo localTimeZone = TimeZoneInfo.FindSystemTimeZoneById(TimeZone.CurrentTimeZone.StandardName);
                TimeZoneInfo eventTimeZone = TimeZoneInfo.FindSystemTimeZoneById(eventTZ);

                String webPageUrl = ConfigurationManager.AppSettings["Url.Authority.Website"].ToString();
                String sData = File.ReadAllText(@"C:\inetpub\wwwroot\EvoCMS.Live\Admin\FOLDERS\Service\Templates\OutlookCalendarAppointment.ics");

                using (SqlConnection cn = new SqlConnection(ConfigurationManager.ConnectionStrings["DataSource.iMIS.Connection"].ToString()))
                {
                    cn.Open();

                    using (SqlCommand cmd = new SqlCommand("select top 1 m.meeting, m.title, m.description, pf.BEGIN_DATE_TIME, pf.END_DATE_TIME, m.address_1, m.address_2, m.address_3, m.city, m.state_province, m.zip, m.country, m.notes, m.contact_id from meet_master m, Product_Function pf, Product p where meeting = @meeting and p.PRODUCT_CODE = pf.PRODUCT_CODE and p.PRODUCT_MAJOR = @meeting and pf.PRODUCT_CODE = @productCode", cn))
                    {
                        cmd.Parameters.AddWithValue("meeting", eventCode);
                        cmd.Parameters.AddWithValue("productCode", productCode);

                        using (SqlDataReader reader = cmd.ExecuteReader(System.Data.CommandBehavior.Default))
                        {
                            while (reader.Read())
                            {
                                // Parse the string data.

                                string venue = (string.IsNullOrEmpty(reader[5].ToString()) ? "" : reader[5].ToString())
                                    + (string.IsNullOrEmpty(reader[6].ToString()) ? "" : (", " + reader[6].ToString()))
                                    + (string.IsNullOrEmpty(reader[7].ToString()) ? "" : (", " + reader[7].ToString()))
                                    + (string.IsNullOrEmpty(reader[8].ToString()) ? "" : (", " + reader[8].ToString()))
                                    + (string.IsNullOrEmpty(reader[9].ToString()) ? "" : (", " + reader[9].ToString()))
                                    + (string.IsNullOrEmpty(reader[10].ToString()) ? "" : (", " + reader[10].ToString()))
                                    + (string.IsNullOrEmpty(reader[11].ToString()) ? "" : (", " + reader[11].ToString()));

                                DateTime start = new DateTime();
                                DateTime end = new DateTime();

                                var startDate = DateTime.TryParse(reader[3].ToString(), out start) ? DateTime.Parse(reader[3].ToString()) : DateTime.Now;
                                var endDate = DateTime.TryParse(reader[4].ToString(), out end) ? DateTime.Parse(reader[4].ToString()) : DateTime.Now;

                                DateTime eventStartTime = TimeZoneInfo.ConvertTimeToUtc(startDate, eventTimeZone);
                                DateTime eventEndTime = TimeZoneInfo.ConvertTimeToUtc(endDate, eventTimeZone);
                                String EventRelativeUrl = String.Format("{0}/{1}/{2}?EventCode={3}",
                                    "/events/eventdetails",
                                    endDate.ToString("yyyy/M"),
                                    eventCode.ToLower(),
                                    eventCode);

                                sData = sData.Replace("__SUBJECT__", "CEDA event | " + reader[1].ToString());
                                sData = sData.Replace("__BODY__", reader[2].ToString());
                                sData = sData.Replace("__TITLE__", reader[1].ToString());
                                sData = sData.Replace("__DESCRIPTION__", "<div style='font-family: Calibri, Helvetica, Arial, Sans-Serif;'>" + reader[2].ToString());
                                sData = sData.Replace("__EVENTURL__", webPageUrl + EventRelativeUrl);
                                sData = sData.Replace("__POLICY__", (reader[12].ToString() + "</div>").Replace("\r", "").Replace("</br>", "<br><br>"));
                                sData = sData.Replace("<br><b>Cancellations", "<b>Cancellations");
                                sData = sData.Replace("please email the event coordinator directly.<BR/><BR/>", "please email the event coordinator directly.<br/><br/><br/>");
                                sData = sData.Replace("__START_DATE__", eventStartTime.ToString("yyyyMMddTHHmm00"));
                                sData = sData.Replace("__END_DATE__", eventEndTime.ToString("yyyyMMddTHHmm00"));
                                sData = sData.Replace("__LOCATION__", venue.Replace("\n", " ").Replace(",", " ").Replace("&nbsp;", " "));
                                sData = sData.Replace("__GUID__", Guid.NewGuid().ToString());

                                eventName = reader[1].ToString();
                            }
                        }
                    }
                }

                using (SqlConnection cn2 = new SqlConnection(ConfigurationManager.ConnectionStrings["DataSource.iMIS.Connection"].ToString()))
                {
                    cn2.Open();

                    using (SqlCommand cmd2 = new SqlCommand("select full_name, work_phone, email from name where id = @id", cn2))
                    {
                        cmd2.Parameters.AddWithValue("id", coordinatorId);

                        using (SqlDataReader reader = cmd2.ExecuteReader(System.Data.CommandBehavior.Default))
                        {
                            while (reader.Read())
                            {
                                // Cater for events coordinator
                                Boolean hasAnEventCoordinator = string.IsNullOrEmpty(coordinatorId) ? false : true;
                                if (hasAnEventCoordinator)
                                {
                                    // Get coordinator contact details
                                    sData = sData.Replace("__COORDINATORNAME__", reader["full_name"].ToString());
                                    sData = sData.Replace("__COORDINATORPHONE__", reader["work_phone"].ToString());
                                    sData = sData.Replace("__COORDINATOREMAIL__", reader["email"].ToString());

                                    coordName = reader["full_name"].ToString();
                                    coordPhone = reader["work_phone"].ToString();
                                    coordEmail = reader["email"].ToString();
                                }
                                else
                                {
                                    sData = sData.Replace("__COORDINATORNAME__", "CEDA");
                                    sData = sData.Replace("__COORDINATORPHONE__", "03 9662 3544");
                                    sData = sData.Replace("__COORDINATOREMAIL__", "info@ceda.com.au");
                                }
                            }
                        }
                    }

                    using (SqlCommand contactCommand = new SqlCommand("select top 1 prefix, full_name from name where id = @contactId", cn2))
                    {
                        contactCommand.Parameters.AddWithValue("contactId", contactId);

                        using (SqlDataReader reader = contactCommand.ExecuteReader(System.Data.CommandBehavior.Default))
                        {
                            while (reader.Read())
                            {
                                contactName = reader["prefix"].ToString() + " " + reader["full_name"].ToString();
                            }
                        }
                    }
                }

                // Create an attachment from the memorystream
                objAttachment = new Attachment(new MemoryStream(new UTF8Encoding(true).GetBytes(sData)), "Add to calendar.ics", "text/calendar; method=REQUEST");
            }
            finally
            {

            }

            return objAttachment;
        }
    }
}
