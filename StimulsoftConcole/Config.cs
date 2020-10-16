using System;
using System.Collections.Generic;
using System.IO;
using System.Web.Script.Serialization;

namespace StimulsoftConsole
{
    public class Config
    {
        public static ReportConfig ReadReportConfig(string configPath)
        {
            string configData;
            ReportConfig ReportConfig = new ReportConfig();

            try
            {
                StreamReader file = new StreamReader(configPath);
                configData = file.ReadToEnd();

                JavaScriptSerializer js = new JavaScriptSerializer();
                ReportConfig confPar = js.Deserialize<ReportConfig>(configData);

                ReportConfig = confPar;

                return ReportConfig;
            }
            catch (Exception)
            {
                ReportConfig.Error = "Ошибка конфигурации файла " + configPath;
                Console.WriteLine("Ошибка конфигурации файла " + configPath);
                return ReportConfig;
            }
        }

        public static EmailConfig ReadEmailConfig(string configPath)
        {
            string configData;
            EmailConfig EmailConfig = new EmailConfig();

            try
            {
                StreamReader file = new StreamReader(configPath);
                configData = file.ReadToEnd();

                JavaScriptSerializer js = new JavaScriptSerializer();
                EmailConfig confPar = js.Deserialize<EmailConfig>(configData);

                EmailConfig = confPar;

                return EmailConfig;
            }
            catch (FileNotFoundException)
            {
                EmailConfig.Error = "Нет конфигурационного файла " + configPath;
                return EmailConfig;
            }
            catch (Exception)
            {
                EmailConfig.Error = "Ошибка конфигурации файла " + configPath;
                Console.WriteLine("Ошибка конфигурации файла " + configPath);
                return EmailConfig;
            }
        }
    }

    public class ReportConfig
    {
        private string exportPath = string.Empty;
        public ProgrammLoyalty[] ProgrammLoyalty { get; set; }
        public string ExportPath
        {
            get { return exportPath; }
            set
            {
                if (!string.IsNullOrEmpty(value))
                    exportPath = value;
            }
        }
        public string Error { get; set; }
    }

    public class EmailConfig
    {
        private int? sendLimit = 1000;
        private int? sendTimeout = 0;
        private bool? send = false;
        public bool? Send {
            get { return send; }
            set
            {
                if (value.HasValue)
                    send = value;
            }
        }
        public int? SendLimit
        {
            get { return sendLimit; }
            set
            {
                if (value.HasValue)
                    sendLimit = value;
            }
        }
        public int? SendTimeout
        {
            get { return sendTimeout; }
            set
            {
                if (value.HasValue)
                    sendTimeout = value;
            }
        }
        public string EmailSender { get; set; }
        public string EmailCopy { get; set; }
        public string Server { get; set; }
        public string Login { get; set; }
        public string Password { get; set; }
        public string Capture { get; set; }
        public string SignatureText { get; set; }
        public string Error { get; set; }
    }

    public class ProgrammLoyalty
    {
        private Reports[] reports = new Reports[0];
        public string DirectoryName { get; set; }
        public string DataSource { get; set; }
        public Reports[] Reports
        {
            get { return reports; }
            set
            {
                if (reports.Length == 0)
                    reports = value;
            }
        }
    }

    public class Reports
    {
        private List<int> legacyCliId = new List<int>();
        private string timeFrom = "09";
        private string timeTo = "09";
        private string divideBy = String.Empty;
        private string byProg = "false";
        public string ReportPath { get; set; }
        public string ReportName { get; set; }
        public string[] Periods { get; set; }
        public string[] ExportFormat { get; set; }
        public string DivideBy
        {
            get { return divideBy; }
            set
            {
                if (string.IsNullOrEmpty(divideBy))
                    divideBy = value.ToLower();
            }
        }
        public string ByProg
        {
            get { return byProg; }
            set
            {
                if (!string.IsNullOrEmpty(byProg))
                    byProg = value;
            }
        }
        public List<int> LegacyCliId
        {
            get { return legacyCliId; }
            set
            {
                if (legacyCliId.Count == 0)
                    legacyCliId = value;
            }
        }
        public string TimeFrom
        {
            get { return timeFrom; }
            set
            {
                if (!string.IsNullOrEmpty(timeFrom))
                    timeFrom = value;
            }
        }
        public string TimeTo
        {
            get { return timeTo; }
            set
            {
                if (!string.IsNullOrEmpty(timeTo))
                    timeTo = value;
            }
        }
    }
}
