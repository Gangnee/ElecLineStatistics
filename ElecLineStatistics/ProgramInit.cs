using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Configuration;
using TestMateAppGroup;
using System.Windows.Forms;

namespace ElecLineStatistics
{
    static class ProgramInit
    {
        /// <summary>
        /// Flag for Program Ruing in IDE or Exe Mode
        /// </summary>
        public static bool IsDevelopMode_g;

        /// <summary>
        /// General Program Information
        /// </summary>
        public static TestMateApp oAppInformation = new TestMateApp("Electronic Line Statistics", "", "", "");

        /// <summary>
        /// General Program Configuration Program
        /// </summary>
        public static Configuration oAppCfg = ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.None);

        public static string strOpDataBase_g;
        public static string strPCBType_g;
        public static string strStatFolder_g;
        public static string strConfirmFolder_g;
        public static string strReportOutput_g;
        public static string strHardIndex_g;
        public static string strStatRoutingDataBase_g;
        public static string strStatusFolder_g;
        public static string strPanelConsumption_g;
        public static string strProductConfirm_g;

        public static string strMonthShiftFolder_g;
        public static string strMonthShiftBackupFolder_g;
        public static string strMonthShiftConfigFolder_g;


        public const string UPDATEPASSWORD_G = "staff2019";
        public const string HARDINDEXPASSWORD_G = "hard2019";

        public const string STOPDATABASEPASSWORD_G = "stop2018";
        public const string SHIFTMODULEPASSWORD_G = "shift2019";

        public const string TEMPFOLDER_G = "D:\\TEMP\\";
        public const string PCBPASSWORD = "elec2018";

        public const string ROUTINGPASSWORD = "routing2018";

        public static string strProgLogFile_g;

        public static string strStopDatabase_g;

        public static string strWinUser_g;
        public static string strMachineName_g;

        public static string strOperatorStatistics_g;

        public static string strStandardRoutingDataBase_g;


        /// <summary>
        /// Porgram History Log Item List
        /// </summary>
        public static void AppHistoryLog()
        {
            oAppInformation.BuildHistory("2020-04-09", "0.01", "nig", "Rebuild UI Setup After Accident from the Elec Statisitics");
            oAppInformation.BuildHistory("2020-04-10", "0.02", "nig", "Build Init Program Structure");
            oAppInformation.BuildHistory("2020-04-11", "0.03", "nig", "ReBuild QD File Converter Process");
            oAppInformation.BuildHistory("2020-04-11", "0.03", "nig", "ReBuild SAP Repair File UpLoad");
            oAppInformation.BuildHistory("2020-04-16", "0.04", "nig", "Routing Management UI Design");
            oAppInformation.BuildHistory("2020-04-17", "0.05", "nig", "Finish Routing Management UI Initialization");
            oAppInformation.BuildHistory("2020-04-18", "0.06", "nig", "Finish Routing Edit Rouging Process");
            oAppInformation.BuildHistory("2020-04-18", "0.06", "nig", "Optimize Structure for Routing Config Process");
            oAppInformation.BuildHistory("2020-04-18", "0.06", "nig", "Add Input Key Sequence");
            oAppInformation.BuildHistory("2020-04-18", "0.06", "nig", "Add Standard Database and Efficience Database File Open");
            oAppInformation.BuildHistory("2020-04-18", "0.06", "nig", "Add Standard Database Initialize During Loading");
        }

        /// <summary>
        /// Validation for Program Config File for Automatic Quit on Update Propose
        /// </summary>
        /// <param name="IsLogUserInfo"></param>
        /// <param name="strAppStatus"></param>
        public static void RunAppValidation(bool IsLogUserInfo = false, string strAppStatus = "Run")
        {
            //Validation Config File to Quit Application
            if (!File.Exists(oAppCfg.FilePath) || (Program.AppErrorHandler.AppException != null && Program.AppErrorHandler.IsErrHandled)) Application.Exit();

            if (IsLogUserInfo)
            {
                string strLogUser = Environment.UserName;
                string strMachineName = Environment.MachineName;
                string strTimer = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");

                string strLogItem = $"[{strTimer}] [{strLogUser }][{strAppStatus}] {strMachineName}";

                TestMateApp.WriteLogFileLine(strLogItem, Application.StartupPath + "\\UserEvent.Log");
            }
        }

        /// <summary>
        /// Applicatino Quit Event for Set Log File, Kill Empty Excel File ...
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        public static void AppQuitEvent(object sender, EventArgs e)
        {
            if (File.Exists(oAppCfg.FilePath))
            {
                RunAppValidation(true, "Quit");
            }
            else
            {
                RunAppValidation(true, "ForceQuit");
            }
        }

        public static void InitGlobals()
        {
            strOpDataBase_g = @oAppCfg.AppSettings.Settings["ElectronicStaffDataBase"].Value;
            strPCBType_g = @oAppCfg.AppSettings.Settings["PCBTypeFile"].Value;
            strStatFolder_g = @oAppCfg.AppSettings.Settings["StatisticsFolder"].Value;
            strConfirmFolder_g = @oAppCfg.AppSettings.Settings["ConfirmFolder"].Value;
            strReportOutput_g = @oAppCfg.AppSettings.Settings["ReportOutput"].Value;
            strHardIndex_g = @oAppCfg.AppSettings.Settings["ProductHardIndex"].Value;
            strStatRoutingDataBase_g = @oAppCfg.AppSettings.Settings["StatRoutingDataBase"].Value;
            strStandardRoutingDataBase_g = @oAppCfg.AppSettings.Settings["StandardRoutingDataBase"].Value;

            strStatusFolder_g = @oAppCfg.AppSettings.Settings["FamilyOutFolder"].Value;
            strPanelConsumption_g = @oAppCfg.AppSettings.Settings["PanelConsumption"].Value;

            strMonthShiftFolder_g = @oAppCfg.AppSettings.Settings["MonthShiftModule"].Value;
            strProductConfirm_g = @oAppCfg.AppSettings.Settings["ProductConfirm"].Value;

            strMonthShiftBackupFolder_g = strMonthShiftFolder_g + "\\InitialBackup\\";
            strMonthShiftConfigFolder_g = strMonthShiftFolder_g + "\\ShiftConfiguration\\";

            strStopDatabase_g = @oAppCfg.AppSettings.Settings["ProdStopDataBase"].Value;
            strOperatorStatistics_g = @oAppCfg.AppSettings.Settings["OperatorStatistics"].Value;

            //Get the Program Log File Name
            strProgLogFile_g = AppDomain.CurrentDomain.BaseDirectory + "\\" + AppDomain.CurrentDomain.SetupInformation.ApplicationName + ".LOG";
            //Get the Environment Information
            strWinUser_g = Environment.UserName;
            strMachineName_g = Environment.MachineName;
        }


    }
}
