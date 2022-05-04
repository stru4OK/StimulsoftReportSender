using System;
using System.Linq;
using System.IO;
using Stimulsoft.Report;
using Stimulsoft.Report.Dictionary;
using System.Collections.Generic;
using System.Reflection;
using System.Threading;

namespace StimulsoftConsole
{
    class Program
    {
        static int Main(string[] args)
        {
            string fileReportConfig = Variables.fileReportConfig;

            string fileEmailConfig = Variables.fileEmailConfig;
            string date = DateTime.Now.ToString("yyyy.MM.dd HH:mm");
            string logFilePath = AppDomain.CurrentDomain.BaseDirectory;

            for (int i = 0; i < args.Length; i += 2)
            {
                switch (args[i])
                {
                    case Variables.flagConfig:
                        fileReportConfig = args[i + 1];
                        break;
                    case Variables.flagEmailConfig:
                        fileEmailConfig = args[i + 1];
                        break;
                    case Variables.flagDate:
                        date = args[i + 1];
                        break;
                    case Variables.flagPath:
                        logFilePath = args[i + 1];

                        if (!Directory.Exists(logFilePath))
                            Directory.CreateDirectory(logFilePath);

                        break;
                    default:
                        Console.WriteLine("Use /config <configFile> or /emailConfig <configFile> or /date <yyyy.MM.dd> or /logsPath <logsPath> flags");
                        return 1;
                }
            }

            if (AdditionalFunc.CheckDate(date) == 1)
            {
                Console.WriteLine("Use /config <configFile> or /emailConfig <configFile> or /date <yyyy.MM.dd> or /logsPath <logsPath> flags");
                return 1;
            }

            Console.WriteLine("StimulsoftConsole v." + Assembly.GetExecutingAssembly().GetName().Version.ToString() + " "
                + Environment.GetEnvironmentVariable("PROCESSOR_ARCHITECTURE").ToString() + " " + DateTime.Now + "\n");
            AdditionalFunc.LogFile(logFilePath + "\\" + Variables.fileLog, "StimulsoftConsole v." + Assembly.GetExecutingAssembly().GetName().Version.ToString() + " "
                + Environment.GetEnvironmentVariable("PROCESSOR_ARCHITECTURE").ToString() + " " + DateTime.Now + "\n", false);

            AdditionalFunc.CheckDriveSpace(logFilePath);

            //date = "2021.05.24";
            //fileReportConfig = "RNiPB_Brs.jsn";

            ReportConfig ReportConfig = Config.ReadReportConfig(fileReportConfig);
            EmailConfig EmailConfig = Config.ReadEmailConfig(fileEmailConfig);

            if (ReportConfig.Error != null)
            {
                AdditionalFunc.LogFile(logFilePath + "\\" + Variables.fileError, ReportConfig.Error + "\n" + EmailConfig.Error, false);
                return 1;
            }

            Reports(logFilePath, ReportConfig, EmailConfig, date);

            return 0;
        }

        public static void Reports(string logFilePath, ReportConfig ReportConfig, EmailConfig EmailConfig, string date)
        {
            List<JuridicalLimErr> JuridicalLimit = new List<JuridicalLimErr>(), JuridicalError = new List<JuridicalLimErr>();

            string jurIDEmailError, jurIDError, jurIDEmailLimit, jurIDLimit, email;
            int index, sendSuccess = 0;

            for (int pl = 0; pl < ReportConfig.ProgrammLoyalty.Length; pl++)
            {
                for (int rp = 0; rp < ReportConfig.ProgrammLoyalty[pl].Reports.Length; rp++)
                {
                    for (int pr = 0; pr < ReportConfig.ProgrammLoyalty[pl].Reports[rp].Periods.Length; pr++)
                    {
                        switch(ReportConfig.ProgrammLoyalty[pl].Reports[rp].DivideBy)
                        {
                            case Variables.divideByJuridical:

                                JuridicalError.Clear();
                                JuridicalLimit.Clear();
                                jurIDEmailError = string.Empty;
                                jurIDError = string.Empty;
                                jurIDEmailLimit = string.Empty;
                                jurIDLimit = string.Empty;

                                List<JuridicalData> JuridicalData = AdditionalFunc.GetJuridicalDataDB(ReportConfig.ProgrammLoyalty[pl].DataSource, logFilePath);

                                if (JuridicalData.Count == 0)
                                {
                                    Console.WriteLine("Were found 0 legacy clients");
                                    break;
                                }

                                for (int i = 0; i < JuridicalData.Count; i++)
                                {
                                    index = ReportConfig.ProgrammLoyalty[pl].Reports[rp].LegacyCliId.IndexOf(JuridicalData[i].jurID);

                                    if (ReportConfig.ProgrammLoyalty[pl].Reports[rp].LegacyCliId.Count == 0 | index >= 0)
                                    {
                                        GenerateReportResult GenerateReport = Report(logFilePath,
                                            ReportConfig.ProgrammLoyalty[pl].DirectoryName,
                                            ReportConfig.ProgrammLoyalty[pl].DataSource,
                                            ReportConfig.ProgrammLoyalty[pl].Reports[rp].ReportPath,
                                            ReportConfig.ExportPath,
                                            ReportConfig.ProgrammLoyalty[pl].Reports[rp].ReportName,
                                            ReportConfig.ProgrammLoyalty[pl].Reports[rp].Periods[pr].ToLower(),
                                            ReportConfig.ProgrammLoyalty[pl].Reports[rp].TimeFrom,
                                            ReportConfig.ProgrammLoyalty[pl].Reports[rp].TimeTo,
                                            ReportConfig.ProgrammLoyalty[pl].Reports[rp].ExportFormat,
                                            ReportConfig.ProgrammLoyalty[pl].Reports[rp].ByProg,
                                            JuridicalData[i],
                                            date);

                                        if (GenerateReport.generateResult == 0)
                                        {
                                            if (EmailConfig.Send.Value)
                                            {
                                                email = JuridicalData[i].email;

                                                if (!string.IsNullOrEmpty(email))
                                                {
                                                    AdditionalFunc.LogFile(logFilePath + "\\" + Variables.fileEmailRouting, "Найдена почта " + email + " по ЮЛ " + JuridicalData[i].jurInnAndTitle + ".", false);

                                                    if (EmailConfig.SendLimit > sendSuccess)
                                                    {
                                                        if (DispatchReport(logFilePath, GenerateReport.ExportFiles, EmailConfig, JuridicalData[i].jurInnAndTitle, email) == 0)
                                                            sendSuccess++;
                                                        else
                                                            JuridicalError.Add(new JuridicalLimErr { jurID = JuridicalData[i].jurID, email = email });

                                                        //таймаут при следующей отправке
                                                        Thread.Sleep((int)EmailConfig.SendTimeout * 1000);
                                                    }
                                                    else
                                                    {
                                                        AdditionalFunc.LogFile(logFilePath + "\\" + Variables.fileEmailRouting, "Превышен лимит на отправку сообщений.", false);

                                                        JuridicalLimit.Add(new JuridicalLimErr { jurID = JuridicalData[i].jurID, email = email });
                                                    }
                                                }
                                                else
                                                    AdditionalFunc.LogFile(logFilePath + "\\" + Variables.fileEmailRouting, "Не найдена почта по ЮЛ " + JuridicalData[i].jurInnAndTitle + ".", false);
                                            }
                                        }
                                        else
                                            JuridicalError.Add(new JuridicalLimErr { jurID = JuridicalData[i].jurID, email = string.Empty });
                                    }
                                }

                                if (JuridicalError.Count != 0)
                                {
                                    for (int i = 0; i < JuridicalError.Count; i++)
                                    {
                                        jurIDEmailError = jurIDEmailError + JuridicalError[i].jurID + " - " + JuridicalError[i].email + "\n";

                                        if (string.IsNullOrEmpty(jurIDError))
                                            jurIDError = JuridicalError[i].jurID.ToString();
                                        else
                                            jurIDError = jurIDError + "," + JuridicalError[i].jurID;
                                    }

                                    AdditionalFunc.LogFile(logFilePath + "\\" + Variables.fileIDEmailError, jurIDEmailError, false);
                                    AdditionalFunc.LogFile(logFilePath + "\\" + Variables.fileIDError, jurIDError, false);
                                }

                                if (JuridicalLimit.Count != 0)
                                {
                                    for (int i = 0; i < JuridicalLimit.Count; i++)
                                    {
                                        jurIDEmailLimit = jurIDEmailLimit + JuridicalLimit[i].jurID + " - " + JuridicalLimit[i].email + "\n";

                                        if (string.IsNullOrEmpty(jurIDLimit))
                                            jurIDLimit = JuridicalLimit[i].jurID.ToString();
                                        else
                                            jurIDLimit = jurIDLimit + "," + JuridicalLimit[i].jurID;
                                    }

                                    AdditionalFunc.LogFile(logFilePath + "\\" + Variables.fileIDEmailLimit, jurIDEmailLimit, false);
                                    AdditionalFunc.LogFile(logFilePath + "\\" + Variables.fileIDLimit, jurIDLimit, false);
                                }
                                break;

                            case Variables.divideByClub:
                                List<BonusClub> BonusClub = AdditionalFunc.GetBonusClubDataDB(ReportConfig.ProgrammLoyalty[pl].DataSource, logFilePath);

                                for (int i = 0; i < BonusClub.Count; i++)
                                {
                                    Report(logFilePath,
                                        ReportConfig.ProgrammLoyalty[pl].DirectoryName,
                                        ReportConfig.ProgrammLoyalty[pl].DataSource,
                                        ReportConfig.ProgrammLoyalty[pl].Reports[rp].ReportPath,
                                        ReportConfig.ExportPath,
                                        ReportConfig.ProgrammLoyalty[pl].Reports[rp].ReportName,
                                        ReportConfig.ProgrammLoyalty[pl].Reports[rp].Periods[pr].ToLower(),
                                        ReportConfig.ProgrammLoyalty[pl].Reports[rp].TimeFrom,
                                        ReportConfig.ProgrammLoyalty[pl].Reports[rp].TimeTo,
                                        ReportConfig.ProgrammLoyalty[pl].Reports[rp].ExportFormat,
                                        ReportConfig.ProgrammLoyalty[pl].Reports[rp].ByProg,
                                        BonusClub[i],
                                        date);
                                }
                                break;

                            default:
                                Report(logFilePath,
                                    ReportConfig.ProgrammLoyalty[pl].DirectoryName,
                                    ReportConfig.ProgrammLoyalty[pl].DataSource,
                                    ReportConfig.ProgrammLoyalty[pl].Reports[rp].ReportPath,
                                    ReportConfig.ExportPath,
                                    ReportConfig.ProgrammLoyalty[pl].Reports[rp].ReportName,
                                    ReportConfig.ProgrammLoyalty[pl].Reports[rp].Periods[pr].ToLower(),
                                    ReportConfig.ProgrammLoyalty[pl].Reports[rp].TimeFrom,
                                    ReportConfig.ProgrammLoyalty[pl].Reports[rp].TimeTo,
                                    ReportConfig.ProgrammLoyalty[pl].Reports[rp].ExportFormat,
                                    ReportConfig.ProgrammLoyalty[pl].Reports[rp].ByProg.ToLower(),
                                    date);
                                break;
                        }
                    }
                }
            }
        }

        public static GenerateReportResult Report(string logPath, string directoryName, string dataSource, string reportPath, string exportPath, string reportName, string reportPeriod, string timeFrom,
            string timeTo, string[] exportFormat, string byProg, string date)
        {
            return Report(logPath, directoryName, dataSource, reportPath, exportPath, reportName, reportPeriod, timeFrom, timeTo, exportFormat, byProg,
                new JuridicalData { jurInnAndTitle = String.Empty, jurID = 0 }, new BonusClub { title = String.Empty, clubId = 0  }, date);
        }

        public static GenerateReportResult Report(string logPath, string directoryName, string dataSource, string reportPath, string exportPath, string reportName, string reportPeriod, string timeFrom,
            string timeTo, string[] exportFormat, string byProg, JuridicalData JuridicalData, string date)
        {
            return Report(logPath, directoryName, dataSource, reportPath, exportPath, reportName, reportPeriod, timeFrom, timeTo, exportFormat, byProg, JuridicalData, 
                new BonusClub { title = String.Empty, clubId = 0 }, date);
        }

        public static GenerateReportResult Report(string logPath, string directoryName, string dataSource, string reportPath, string exportPath, string reportName, string reportPeriod, string timeFrom,
            string timeTo, string[] exportFormat, string byProg, BonusClub BonusClub, string date)
        {
            return Report(logPath, directoryName, dataSource, reportPath, exportPath, reportName, reportPeriod, timeFrom, timeTo, exportFormat, byProg, new JuridicalData { jurInnAndTitle = String.Empty, jurID = 0 },
                BonusClub, date);
        }

        public static GenerateReportResult Report(string logFilePath, string directoryName, string dataSource, string reportPath, string exportPath, string reportName, string reportPeriod, string timeFrom,
            string timeTo, string[] exportFormat, string byProg, JuridicalData JuridicalData, BonusClub BonusClub, string date)
        {
            GenerateReportResult GenerateReportResult = GenerateReport(logFilePath, directoryName, dataSource, reportPath, AppDomain.CurrentDomain.BaseDirectory + "\\" + exportPath + "\\" + reportName, reportPeriod, 
                timeFrom, timeTo, exportFormat, byProg, JuridicalData, BonusClub, date);

            if (!string.IsNullOrEmpty(JuridicalData.jurInnAndTitle) & JuridicalData.jurID != 0)
            {
                if (GenerateReportResult.generateResult == 0)
                {
                    Console.WriteLine(DateTime.Now.ToString("HH:mm:ss") + " В программе " + directoryName + " " + reportName + " за " + reportPeriod + " для ЮЛ "
                        + JuridicalData.jurInnAndTitle + " сформирован!");
                    AdditionalFunc.LogFile(logFilePath + "\\" + Variables.fileLog, "В программе " + directoryName + " отчет " + reportName + " за " + reportPeriod + " для ЮЛ "
                        + JuridicalData.jurInnAndTitle + " сформирован!", false);
                }
                else
                {
                    Console.WriteLine(DateTime.Now.ToString("HH:mm:ss") + " В программе " + directoryName + " " + reportName + " за " + reportPeriod + " для ЮЛ "
                        + JuridicalData.jurInnAndTitle + " НЕ сформирован!");
                    AdditionalFunc.LogFile(logFilePath + "\\" + Variables.fileLog, "В программе " + directoryName + " отчет " + reportName + " за " + reportPeriod + " для ЮЛ "
                        + JuridicalData.jurInnAndTitle + " НЕ сформирован!", false);
                }
            }
            else if(!string.IsNullOrEmpty(BonusClub.title) & BonusClub.clubId != 0)
            {
                if (GenerateReportResult.generateResult == 0)
                {
                    Console.WriteLine(DateTime.Now.ToString("HH:mm:ss") + " В программе " + BonusClub.title + " " + reportName + " за " + reportPeriod + " сформирован!");
                    AdditionalFunc.LogFile(logFilePath + "\\" + Variables.fileLog, "В программе " + BonusClub.title + " отчет " + reportName + " за " + reportPeriod + " сформирован!", false);
                }
                else
                {
                    Console.WriteLine(DateTime.Now.ToString("HH:mm:ss") + " В программе " + directoryName + " " + reportName + " за " + reportPeriod + " НЕ сформирован!");
                    AdditionalFunc.LogFile(logFilePath + "\\" + Variables.fileLog, "В программе " + directoryName + " отчет " + reportName + " за " + reportPeriod + " НЕ сформирован!", false);
                }
            }
            else
            {
                if (GenerateReportResult.generateResult == 0)
                {
                    Console.WriteLine(DateTime.Now.ToString("HH:mm:ss") + " В программе " + directoryName + " " + reportName + " за " + reportPeriod + " сформирован!");
                    AdditionalFunc.LogFile(logFilePath + "\\" + Variables.fileLog, "В программе " + directoryName + " отчет " + reportName + " за " + reportPeriod + " сформирован!", false);
                }
                else
                {
                    Console.WriteLine(DateTime.Now.ToString("HH:mm:ss") + " В программе " + directoryName + " " + reportName + " за " + reportPeriod + " НЕ сформирован!");
                    AdditionalFunc.LogFile(logFilePath + "\\" + Variables.fileLog, "В программе " + directoryName + " отчет " + reportName + " за " + reportPeriod + " НЕ сформирован!", false);
                }
            }

            return GenerateReportResult;
        }

        public static GenerateReportResult GenerateReport(string logFilePath, string directoryName, string dataSource, string reportPath, string reportName, string reportPeriod,
            string timeFrom, string timeTo, string[] exportFormat, string byProg, JuridicalData JuridicalData, BonusClub BonusClub, string date)
        {
            StiExportFormat StiExportFormat = StiExportFormat.Pdf;
            GenerateReportResult GenerateReportResult = new GenerateReportResult();
            List<string> ExportFiles = new List<string>();

            DateTime monday = AdditionalFunc.GetMonday(date);
            string numWeek = AdditionalFunc.GetNumWeek(date), exportPath = string.Empty, exportFileName = string.Empty, bonusClubFolder = "\\" + BonusClub.title;

            const string period = "period", periodRequest = "periodRequest", timeReport = "timeReport", 
                juridicalID = "juridicalID", clubID = "clubID", shiftHour = "shiftHour", programDivide = "programDivide", 
                year = "year", month = "month"; 

            StiReport report = new StiReport();

            try
            {
                StiConfig.Services.Add(new StiOracleODPAdapterService());

                StiOptions.Engine.ReportCache.AmountOfQuickAccessPages = 5;
                StiOptions.Engine.ReportCache.AmountOfProcessedPagesForStartGCCollect = 5;
                report.ReportCacheMode = StiReportCacheMode.On;

                report.Load(reportPath);

                foreach (StiOracleODPDatabase db in report.Dictionary.Databases.OfType<StiOracleODPDatabase>())
                    db.ConnectionString = dataSource;

                foreach (StiOracleODPSource src in report.Dictionary.DataSources.OfType<StiOracleODPSource>())
                    src.CommandTimeout = 12000;

                if (report.Dictionary.Variables != null && report.Dictionary.Variables.Contains(timeReport) && report.Dictionary.Variables.Contains(period)
                    && report.Dictionary.Variables.Contains(periodRequest))
                {
                    if (report.Dictionary.Variables.Contains(juridicalID))
                        report.Dictionary.Variables[juridicalID].Value = JuridicalData.jurID.ToString();

                    if (report.Dictionary.Variables.Contains(clubID))
                        report.Dictionary.Variables[clubID].Value = BonusClub.clubId.ToString();

                    if (report.Dictionary.Variables.Contains(shiftHour))
                        report.Dictionary.Variables[shiftHour].Value = timeFrom;

                    if (report.Dictionary.Variables.Contains(programDivide))
                        report.Dictionary.Variables[programDivide].Value = byProg;

                    if (report.Dictionary.Variables.Contains(year))
                        report.Dictionary.Variables[year].Value = date.Substring(0, 4);

                    if (report.Dictionary.Variables.Contains(month))
                        report.Dictionary.Variables[month].Value = date.Substring(5, 2);

                    report.Dictionary.Variables[timeReport].Value = date;

                    switch (reportPeriod)
                    {
                        case Variables.periodAllTime:
                            report.Dictionary.Variables[period].Value = "За все время. Смена с " + timeFrom + ":00 до " + timeTo + ":00";
                            /*report.Dictionary.Variables[periodRequest].Value = "and to_char(request_date, 'yyyy.mm.dd hh24:mi:ss')<to_char(to_date('"
                                + date + "', 'yyyy.mm.dd hh24:mi:ss'), 'yyyy.mm.dd hh24:mi:ss')";*/
                            report.Dictionary.Variables[periodRequest].Value = "and request_date < to_date('" + date + "', 'yyyy.mm.dd hh24:mi:ss') ";
                            //    + "and (((to_char(request_date,'HH24') >= " + timeFrom + " or to_char(request_date,'HH24') < " + timeTo + ") and " + timeFrom + " >= " + timeTo + ") "
                            //    + "or ((to_char(request_date, 'HH24') >= " + timeFrom + " and to_char(request_date, 'HH24') < " + timeTo + ") and " + timeFrom + " < " + timeTo + "))";
                            break;
                        case Variables.periodYear:
                            report.Dictionary.Variables[period].Value = "С 01.01." + DateTime.Parse(date).AddDays(-1).ToString("yyyy") + " " + timeFrom + ":00 по "
                                + DateTime.Parse(date).ToString("dd.MM.yyyy") + " " + timeTo + ":00. Смена с " + timeFrom + ":00 до " + timeTo + ":00";
                            /*report.Dictionary.Variables[periodRequest].Value = "and to_char(request_date, 'yyyy.mm.dd hh24:mi:ss') between to_char(to_date('"
                                + date + "', 'yyyy.mm.dd hh24:mi:ss') - 1, 'yyyy')|| '.01.01 " + timeFrom + ":00:00' and to_char(to_date('" + date + "', 'yyyy.mm.dd hh24:mi:ss'), 'yyyy.mm.dd')|| ' " + timeTo + ":00:00'";*/
                            report.Dictionary.Variables[periodRequest].Value = "and request_date between to_date('" + DateTime.Parse(date).AddDays(-1).ToString("yyyy")
                                + ".01.01 " + timeFrom + ":00:00', 'yyyy.mm.dd hh24:mi:ss') and to_date('" + DateTime.Parse(date).ToString("yyyy.MM.dd") + " " + timeTo + ":00:00', 'yyyy.mm.dd hh24:mi:ss') ";
                            //    + "and (((to_char(request_date,'HH24') >= " + timeFrom + " or to_char(request_date,'HH24') < " + timeTo + ") and " + timeFrom + " >= " + timeTo + ") "
                            //    + "or ((to_char(request_date, 'HH24') >= " + timeFrom + " and to_char(request_date, 'HH24') < " + timeTo + ") and " + timeFrom + " < " + timeTo + "))";
                            break;
                        case Variables.periodMonth:
                            report.Dictionary.Variables[period].Value = "С 01." + DateTime.Parse(date).AddDays(-1).ToString("MM.yyyy") + " " + timeFrom + ":00 по "
                                + DateTime.Parse(date).ToString("dd.MM.yyyy") + " " + timeTo + ":00. Смена с " + timeFrom + ":00 до " + timeTo + ":00";
                            /*report.Dictionary.Variables[periodRequest].Value = "and to_char(request_date, 'yyyy.mm.dd hh24:mi:ss') between to_char(to_date('"
                                + date + "', 'yyyy.mm.dd hh24:mi:ss') - 1, 'yyyy.mm')|| '.01 " + timeFrom + ":00:00' and to_char(to_date('" + date + "', 'yyyy.mm.dd hh24:mi:ss'), 'yyyy.mm.dd')|| ' " + timeTo + ":00:00'";*/
                            report.Dictionary.Variables[periodRequest].Value = "and request_date between to_date('" + DateTime.Parse(date).AddDays(-1).ToString("yyyy.MM") + ".01 " + timeFrom + ":00:00', 'yyyy.mm.dd hh24:mi:ss') "
                                + "and to_date('" + DateTime.Parse(date).ToString("yyyy.MM.dd") + " " + timeTo + ":00:00', 'yyyy.mm.dd hh24:mi:ss') ";
                            //    + "and (((to_char(request_date,'HH24') >= " + timeFrom + " or to_char(request_date,'HH24') < " + timeTo + ") and " + timeFrom + " >= " + timeTo + ") "
                            //    + "or ((to_char(request_date, 'HH24') >= " + timeFrom + " and to_char(request_date, 'HH24') < " + timeTo + ") and " + timeFrom + " < " + timeTo + "))";
                            break;
                        case Variables.periodWeek:
                            report.Dictionary.Variables[period].Value = "С " + monday.ToString("dd.MM.yyyy") + " " + timeFrom + ":00 по " 
                                + DateTime.Parse(date).ToString("dd.MM.yyyy") + " " + timeTo + ":00. Смена с " + timeFrom + ":00 до " + timeTo + ":00";
                            /*report.Dictionary.Variables[periodRequest].Value = "and to_char(request_date, 'yyyy.mm.dd hh24:mi:ss') between to_char(next_day(to_date('"
                                + date + "', 'yyyy.mm.dd hh24:mi:ss') - 8, 1), 'yyyy.mm.dd')|| ' " + timeFrom + ":00:00' and to_char(to_date('" + date + "', 'yyyy.mm.dd hh24:mi:ss'), 'yyyy.mm.dd')|| ' " + timeTo + ":00:00'";*/
                            report.Dictionary.Variables[periodRequest].Value = "and request_date between to_date('" + monday.ToString("yyyy.MM.dd") + " " + timeFrom + ":00:00', 'yyyy.mm.dd hh24:mi:ss') "
                                + "and to_date('" + DateTime.Parse(date).ToString("yyyy.MM.dd") + " " + timeTo + ":00:00', 'yyyy.mm.dd hh24:mi:ss') ";
                            //    + "and (((to_char(request_date,'HH24') >= " + timeFrom + " or to_char(request_date,'HH24') < " + timeTo + ") and " + timeFrom + " >= " + timeTo + ") "
                            //    + "or ((to_char(request_date, 'HH24') >= " + timeFrom + " and to_char(request_date, 'HH24') < " + timeTo + ") and " + timeFrom + " < " + timeTo + "))";
                            break;
                        case Variables.periodDay:
                            report.Dictionary.Variables[period].Value = "С " + DateTime.Parse(date).AddDays(-1).ToString("dd.MM.yyyy") + " " + timeFrom + ":00 по " + DateTime.Parse(date).ToString("dd.MM.yyyy")
                                + " " + timeTo + ":00. Смена с " + timeFrom + ":00 до " + timeTo + ":00";
                            /*report.Dictionary.Variables[periodRequest].Value = "and to_char(request_date, 'yyyy.mm.dd hh24:mi:ss') between to_char(to_date('"
                                + date + "', 'yyyy.mm.dd hh24:mi:ss') - 1, 'yyyy.mm.dd')|| ' " + timeFrom + ":00:00' and to_char(to_date('" + date + "', 'yyyy.mm.dd hh24:mi:ss'), 'yyyy.mm.dd')|| ' " + timeTo + ":00:00'";*/
                            report.Dictionary.Variables[periodRequest].Value = "and request_date between to_date('" + DateTime.Parse(date).AddDays(-1).ToString("yyyy.MM.dd") + " " + timeFrom + ":00:00', 'yyyy.mm.dd hh24:mi:ss') "
                                + "and to_date('" + DateTime.Parse(date).ToString("yyyy.MM.dd") + " " + timeTo + ":00:00', 'yyyy.mm.dd hh24:mi:ss') ";
                            //    + "and (((to_char(request_date,'HH24') >= " + timeFrom + " or to_char(request_date,'HH24') < " + timeTo + ") and " + timeFrom + " >= " + timeTo + ") "
                            //    + "or ((to_char(request_date, 'HH24') >= " + timeFrom + " and to_char(request_date, 'HH24') < " + timeTo + ") and " + timeFrom + " < " + timeTo + "))";
                            break;
                        default:
                            break;
                    }
                }

                report.Save(reportPath);
                report.Render();

                StiOptions.Export.Csv.ForcedSeparator = ";";

                for (int i = 0; i < exportFormat.Length; i++)
                {
                    string ef = exportFormat[i].ToLower();

                    switch (exportFormat[i].ToLower())
                    {
                        case Variables.formatPDF:
                            StiExportFormat = StiExportFormat.Pdf;
                            break;
                        case Variables.formatCSV:
                            StiExportFormat = StiExportFormat.Csv;
                            break;
                        case Variables.formatDOCX:
                            StiExportFormat = StiExportFormat.Word2007;
                            break;
                        case Variables.formatXLSX:
                            StiExportFormat = StiExportFormat.Excel2007;
                            break;
                        default:
                            break;
                    }

                    switch (reportPeriod)
                    {
                        case Variables.periodAllTime:

                            exportPath = reportName + "\\" + directoryName + bonusClubFolder;

                            if (!String.IsNullOrEmpty(JuridicalData.jurInnAndTitle))
                                exportFileName = JuridicalData.jurInnAndTitle + "." + ef;
                            else
                                exportFileName = "За все время-" + directoryName + "." + ef;

                            if (!Directory.Exists(exportPath))
                                Directory.CreateDirectory(exportPath);

                            report.ExportDocument(StiExportFormat, exportPath + "\\" + exportFileName);
                            break;

                        case Variables.periodYear:

                            exportPath = reportName + "\\" + directoryName + bonusClubFolder + "\\" + DateTime.Parse(date).AddDays(-1).ToString("yyyy") + "-" + directoryName;

                            if (!String.IsNullOrEmpty(JuridicalData.jurInnAndTitle))
                                exportFileName = JuridicalData.jurInnAndTitle + "." + ef;
                            else
                                exportFileName = DateTime.Parse(date).AddDays(-1).ToString("yyyy") + "-" + directoryName + "." + ef;

                            if (!Directory.Exists(exportPath))
                                Directory.CreateDirectory(exportPath);

                            report.ExportDocument(StiExportFormat, exportPath + "\\" + exportFileName);
                            break;

                        case Variables.periodMonth:

                            exportPath = reportName + "\\" + directoryName + bonusClubFolder + "\\" + DateTime.Parse(date).AddDays(-1).ToString("yyyy") + "-" + directoryName
                                + "\\" + DateTime.Parse(date).AddDays(-1).ToString("yyyy") + "-" + DateTime.Parse(date).AddDays(-1).ToString("MM") + "-" + directoryName;

                            if (!String.IsNullOrEmpty(JuridicalData.jurInnAndTitle))
                                exportFileName = JuridicalData.jurInnAndTitle + "." + ef;
                            else
                                exportFileName = DateTime.Parse(date).AddDays(-1).ToString("yyyy") + "-" + DateTime.Parse(date).AddDays(-1).ToString("MM")
                                    + "-" + directoryName + "." + ef;
                                
                            if (!Directory.Exists(exportPath))
                                Directory.CreateDirectory(exportPath);

                            report.ExportDocument(StiExportFormat, exportPath + "\\" + exportFileName);
                            break;

                        case Variables.periodWeek:

                            exportPath = reportName + "\\" + directoryName + bonusClubFolder + "\\" + AdditionalFunc.GetYearByNumWeek(numWeek, date) + "-" + directoryName + "\\Недели\\" 
                                + numWeek + "." + monday.ToString("yyyy") + "-" + monday.ToString("MM") + "-" + monday.ToString("dd") + " - " + monday.AddDays(6).ToString("yyyy") + "-" 
                                + monday.AddDays(6).ToString("MM") + "-" + monday.AddDays(6).ToString("dd") + "-" + directoryName;

                            if (!String.IsNullOrEmpty(JuridicalData.jurInnAndTitle))
                                exportFileName = JuridicalData.jurInnAndTitle + "." + ef;
                            else
                                exportFileName = monday.ToString("yyyy") + "-" + monday.ToString("MM") + "-" + monday.ToString("dd") + " - "
                                   + monday.AddDays(6).ToString("yyyy") + "-" + monday.AddDays(6).ToString("MM") + "-" + monday.AddDays(6).ToString("dd") + "-" + directoryName + "." + ef;

                            if (!Directory.Exists(exportPath))
                                Directory.CreateDirectory(exportPath);

                            report.ExportDocument(StiExportFormat, exportPath + "\\" + exportFileName);
                            break;

                        case Variables.periodDay:

                            exportPath = reportName + "\\" + directoryName + bonusClubFolder + "\\" + DateTime.Parse(date).AddDays(-1).ToString("yyyy") + "-" + directoryName
                                + "\\" + DateTime.Parse(date).AddDays(-1).ToString("yyyy") + "-" + DateTime.Parse(date).AddDays(-1).ToString("MM") + "-" + directoryName
                                + "\\" + DateTime.Parse(date).AddDays(-1).ToString("yyyy") + "-" + DateTime.Parse(date).AddDays(-1).ToString("MM") + "-" + DateTime.Parse(date).AddDays(-1).ToString("dd")
                                + "-" + directoryName;

                            if (!String.IsNullOrEmpty(JuridicalData.jurInnAndTitle))
                                exportFileName = JuridicalData.jurInnAndTitle + "." + ef;
                            else
                                exportFileName = DateTime.Parse(date).AddDays(-1).ToString("yyyy") + "-" + DateTime.Parse(date).AddDays(-1).ToString("MM")
                                + "-" + DateTime.Parse(date).AddDays(-1).ToString("dd") + "-" + directoryName + "." + ef;

                            if (!Directory.Exists(exportPath))
                                Directory.CreateDirectory(exportPath);

                            report.ExportDocument(StiExportFormat, exportPath + "\\" + exportFileName);
                            break;

                        default:
                            break;
                    }

                    ExportFiles.Add(exportPath + "\\" + exportFileName);
                    GenerateReportResult.ExportFiles = ExportFiles;
                }

                GenerateReportResult.generateResult = 0;

                return GenerateReportResult;
            }
            catch (Exception ex)
            {
                AdditionalFunc.LogFile(logFilePath + "\\" + Variables.fileError, "Отчет: " + reportPath + "\n\n" + "Период: " + reportPeriod + "\n\n" + "ЮЛ: " + JuridicalData.jurInnAndTitle 
                    + " " + JuridicalData.jurID + "\n\n" + "БД: " + dataSource.Remove(dataSource.IndexOf(";")) + "\n\n" + ex.Message, true);
                GenerateReportResult.generateResult = 1;

                 return GenerateReportResult;
            }
            finally
            {
                report.Dictionary.Clear();
                report.RenderedPages.Clear();
                report.Dispose();
            }
        }

        public static int DispatchReport(string logPath, List<string> exportFiles, EmailConfig EmailConfig, string jurInnAndTitle, string email)
        {
            if (AdditionalFunc.SendMail(logPath + "\\" + Variables.fileError, EmailConfig.Server, EmailConfig.EmailSender, EmailConfig.Login, EmailConfig.Password, email, EmailConfig.EmailCopy,
                EmailConfig.Capture + " для " + jurInnAndTitle.Remove(0, jurInnAndTitle.IndexOf("-") + 1), EmailConfig.SignatureText, exportFiles) == 0)
            {
                AdditionalFunc.LogFile(logPath + "\\" + Variables.fileEmailRouting, "Письмо с отчетами адресату " + email + " (" + jurInnAndTitle + ") успешно отправлено!", false);
                return 0;
            }
            else
            {
                AdditionalFunc.LogFile(logPath + "\\" + Variables.fileEmailRouting, "Письмо с отчетами адресату " + email + " (" + jurInnAndTitle + ") НЕ отправлено!", false);
                return 1;
            }
        }
    }
}
