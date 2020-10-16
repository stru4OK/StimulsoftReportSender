using Oracle.ManagedDataAccess.Client;
using System;
using System.Collections.Generic;
using System.IO;
using System.Net;
using System.Net.Mail;

namespace StimulsoftConsole
{
    public class AdditionalFunc
    {
        public static string DataBaseSQL(string oracleDbConnection, string sql)
        {
            string data = String.Empty;

            OracleConnection conn = new OracleConnection(oracleDbConnection);

            try
            {
                conn.Open();
                OracleCommand dbcmd = conn.CreateCommand();
                dbcmd.CommandText = sql;
                OracleDataReader reader = dbcmd.ExecuteReader();

                while (reader.Read())
                {
                    data = (string)reader["data"].ToString();
                }

                reader.Close();
                reader = null;
                
                return data;
            }

            catch (Exception ex)
            {
                LogFile(Variables.fileError, oracleDbConnection.Remove(oracleDbConnection.IndexOf(";")) + "\n\n" + ex.ToString(), false);
                return data;
            }
            finally
            {
                OracleConnection.ClearPool(conn);
                conn.Dispose();
                conn.Close();
                conn = null;
            }
        }

        public static List<JuridicalData> GetJuridicalDataDB(string oracleDBConnection)
        {
            List<JuridicalData> JD = new List<JuridicalData>();

            OracleConnection conn = new OracleConnection(oracleDBConnection);

            try
            {
                conn.Open();
                OracleCommand dbcmd = conn.CreateCommand();

                dbcmd.CommandText = "select replace((case when inn is not null then inn ||'-'|| (case when lc.title is not null then lc.title else rp.title end) "
                    + "else (case when lc.title is not null then lc.title else rp.title end) end), '\"') as jurInnAndTitle, "
                    + "lc.inn, (case when lc.title is not null then lc.title else rp.title end) as jurTitle, lc.legacy_cli_id as jurID, "
                    + "lc.email from legacy_clients lc join retail_points rp on lc.legacy_cli_id = rp.legacy_cli_id "
                    + "where lower(lc.agent_agreement) not like '%(раст)%' "
                    + "group by lc.inn, (case when lc.title is not null then lc.title else rp.title end), lc.legacy_cli_id, lc.email order by 3";

                OracleDataReader reader = dbcmd.ExecuteReader();
                while (reader.HasRows)
                {
                    while (reader.Read())
                    {
                        JD.Add(new JuridicalData()
                        {
                            jurInnAndTitle = reader.IsDBNull(0) ? String.Empty : reader.GetString(0),
                            inn = reader.IsDBNull(1) ? String.Empty : reader.GetString(1),
                            jurTitle = reader.IsDBNull(2) ? String.Empty : reader.GetString(2),
                            jurID = reader.IsDBNull(3) ? 0 : reader.GetInt32(3),
                            email = reader.IsDBNull(4) ? String.Empty : reader.GetString(4)
                        });
                    }

                    reader.NextResult();
                }
                reader.Close();
                reader = null;

                return JD;
            }

            catch (Exception ex)
            {
                Console.WriteLine(DateTime.Now + ":\n" + ex.ToString() + "\n\n");
                return JD;
            }
            finally
            {
                OracleConnection.ClearPool(conn);
                conn.Dispose();
                conn.Close();
                conn = null;
            }
        }

        public static List<BonusClub> GetBonusClubDataDB(string oracleDBConnection)
        {
            List<BonusClub> BC = new List<BonusClub>();

            OracleConnection conn = new OracleConnection(oracleDBConnection);

            try
            {
                conn.Open();
                OracleCommand dbcmd = conn.CreateCommand();

                dbcmd.CommandText = "select club_id, title from bonus_clubs order by 2";

                OracleDataReader reader = dbcmd.ExecuteReader();
                while (reader.HasRows)
                {
                    while (reader.Read())
                    {
                        BC.Add(new BonusClub()
                        {
                            clubId = reader.IsDBNull(0) ? 0 : reader.GetInt32(0),
                            title = reader.IsDBNull(1) ? String.Empty : reader.GetString(1)
                        });
                    }

                    reader.NextResult();
                }
                reader.Close();
                reader = null;

                return BC;
            }

            catch (Exception ex)
            {
                Console.WriteLine(DateTime.Now + ":\n" + ex.ToString() + "\n\n");
                return BC;
            }
            finally
            {
                OracleConnection.ClearPool(conn);
                conn.Dispose();
                conn.Close();
                conn = null;
            }
        }

        public static int LogFile(string logFilePath, string text, bool sendEmail)
        {
            if (sendEmail)
                SendMail(logFilePath, Variables.emailSMTP, Variables.emailSender, Variables.emailLogin, Variables.emailPassword, "helpdesk@groupbms.ru", string.Empty, "Ошибка при автоматической ежепериодной выгрузке отчета",
                    Environment.MachineName + " : " + AppDomain.CurrentDomain.BaseDirectory + "\n\n" + text, new List<string>());

            StreamWriter files = new StreamWriter(logFilePath, true);

            files.Write(DateTime.Now.ToString("yyyy.MM.dd HH:mm:ss") + " : \n" + text + "\n\n");

            files.Flush();
            files.Close();
            return 0;
        }

        public static int CheckDriveSpace(string logsPath)
        {
            try
            {
                FileInfo file = new FileInfo(AppDomain.CurrentDomain.BaseDirectory);
                DriveInfo d = new DriveInfo(file.Directory.Root.FullName);

                if (d.IsReady == true)
                    if (d.AvailableFreeSpace * 100 / d.TotalSize < 10)
                    {
                        LogFile(logsPath + "\\" + Variables.fileError, "Свободного дискового пространства осталось меньше 10%", true);
                        return 1;
                    }
                return 0;
            }
            catch (Exception ex)
            {
                LogFile(logsPath + "\\" + Variables.fileError, ex.ToString(), true);
                return 1;
            }
        }

        public static int SendMail(string logsPath, string server, string emailSender, string login, string password, 
            string mailTo, string emailCopy, string caption, string signatureText, List<string> attachFiles)
        {
            try
            {
                MailMessage mail = new MailMessage();

                if (!String.IsNullOrEmpty(emailSender))
                {
                    mail.ReplyToList.Add(new MailAddress(emailSender, "reply-to"));
                    mail.From = new MailAddress(emailSender, emailSender);
                }
                else
                    mail.From = new MailAddress(login);

                if (!String.IsNullOrEmpty(emailCopy))
                {
                    string[] emailCopyArray = emailCopy.Split(';');

                    foreach (string emailCopyElement in emailCopyArray)
                    {
                        mail.CC.Add(new MailAddress(emailCopyElement));
                    }
                }

                if (!String.IsNullOrEmpty(mailTo))
                {
                    string[] emailToArray = mailTo.Split(';');

                    foreach (string emailToElement in emailToArray)
                    {
                        mail.To.Add(new MailAddress(emailToElement));
                    }
                }

                mail.Subject = caption;
                mail.Body = signatureText;

                for (int i = 0; i < attachFiles.Count; i++)
                    if (!string.IsNullOrEmpty(attachFiles[i]))
                        mail.Attachments.Add(new Attachment(attachFiles[i]));

                SmtpClient client = new SmtpClient();
                client.Host = server;
                client.Port = 587;
                client.EnableSsl = true;

                ServicePointManager.ServerCertificateValidationCallback += delegate (object sender, System.Security.Cryptography.X509Certificates.X509Certificate certificate,
                        System.Security.Cryptography.X509Certificates.X509Chain chain,
                        System.Net.Security.SslPolicyErrors sslPolicyErrors)
                {
                    return true;
                };

                client.Timeout = 10000;
                client.UseDefaultCredentials = false;
                client.Credentials = new NetworkCredential(login, password);
                client.DeliveryMethod = SmtpDeliveryMethod.Network;
                client.Send(mail);
                mail.Dispose();
                return 0;
            }
            catch (Exception ex)
            {
                LogFile(logsPath + "\\" + Variables.fileError, ex.ToString(), false);
                return 1;
            }
        }

        public static DateTime GetMonday(string date)
        {
            DateTime Date = DateTime.Parse(date);

            if (Date.DayOfWeek == DayOfWeek.Monday)
                Date = Date.AddDays(-7);

            while (Date.DayOfWeek != DayOfWeek.Monday)
                Date = Date.AddDays(-1);

            return Date;
        }

        public static string GetNumWeek(string date)
        {
            /*CultureInfo cult = new CultureInfo("ru-RU");
            int week = cult.Calendar.GetWeekOfYear(DateTime.Parse(date), CalendarWeekRule.FirstFourDayWeek, DayOfWeek.Monday);
            Console.WriteLine(week);*/

            DateTime Date = DateTime.Parse(date).AddDays(-1);

            int a = (14 - Convert.ToInt32(Date.ToString("MM"))) / 12;
            int y = Convert.ToInt32(Date.ToString("yyyy")) + 4800 - a;
            int m = Convert.ToInt32(Date.ToString("MM")) + 12 * a - 3;
            int JD = Convert.ToInt32(Date.ToString("dd")) + (153 * m + 2) / 5 + 365 * y + y / 4 - y / 100 + y / 400 - 32045;
            int d4 = (JD + 31741 - (JD % 7)) % 146097 % 36524 % 1461;
            int L = d4 / 1460;
            int d1 = ((d4 - L) % 365) + L;
            int WN = (d1 / 7) + 1;

            return WN.ToString("00");
        }

        public static string GetYearByNumWeek(string numWeek, string date)
        {
            if (DateTime.Parse(date).AddDays(-1).ToString("MM") == "01" & Convert.ToInt16(numWeek) > 50)
                return DateTime.Parse(date).AddDays(-1).AddYears(-1).ToString("yyyy");
            if (DateTime.Parse(date).AddDays(-1).ToString("MM") == "12" & Convert.ToInt16(numWeek) > 50)
                return DateTime.Parse(date).AddDays(-1).ToString("yyyy");
            if (DateTime.Parse(date).ToString("MM") == "12" & Convert.ToInt16(numWeek) < 5)
                return DateTime.Parse(date).AddDays(-1).AddYears(1).ToString("yyyy");
            else
                return DateTime.Parse(date).ToString("yyyy");
        }

        public static int CheckDate(string date)
        {
            try
            {
                DateTime.Parse(date);
                return 0;
            }
            catch (Exception)
            {
                return 1;
            }
        }
    }
}
