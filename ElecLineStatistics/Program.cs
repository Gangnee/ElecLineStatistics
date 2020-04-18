using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using TestMateAppGroup;
using WinFormExcel;


namespace ElecLineStatistics
{

    static class Program
    {

        /// <summary>
        /// Error Handler for Program
        /// </summary>
        public static TMErrorHandler AppErrorHandler = new TMErrorHandler(ProgramInit.oAppInformation);

        /// <summary>
        /// 应用程序的主入口点。
        /// </summary>
        [STAThread]
        static void Main()
        {
            //Flag for the Environemnt of the Application
            ProgramInit.IsDevelopMode_g = false;
            #if DEBUG
            ProgramInit.IsDevelopMode_g = true;
            #endif

            ProgramInit.AppHistoryLog();

            if (!ProgramInit.IsDevelopMode_g)
            {
                try
                {
                    //Init Globals Information from Configuration File
                    ProgramInit.InitGlobals();
                    Application.SetUnhandledExceptionMode(UnhandledExceptionMode.CatchException);
                    Application.ThreadException += new System.Threading.ThreadExceptionEventHandler(AppErrorHandler.Application_ThreadException);
                    AppDomain.CurrentDomain.UnhandledException += new UnhandledExceptionEventHandler(AppErrorHandler.CurrentDomain_UnhandledException);

                    //Add the Quit Event on Application
                    Application.ApplicationExit += ProgramInit.AppQuitEvent;
                    //Cleanup the BackGround Excel Application
                    //Application.Idle += MsExcelFile.EventCleanupEmptyBGExcel;

                    Application.EnableVisualStyles();
                    Application.SetCompatibleTextRenderingDefault(false);
                    Application.Run(new frmMain());
                }
                catch (Exception ex)
                {
                    AppErrorHandler.PopupException(ex, true);
                    Application.Exit();
                }
            }
            else
            {
                //Init Globals Information from Configuration File
                ProgramInit.InitGlobals();

                //Add tge Quit Event
                Application.ApplicationExit += ProgramInit.AppQuitEvent;
                Application.Idle += MsExcelFile.EventCleanupEmptyBGExcel;

                Application.EnableVisualStyles();
                Application.SetCompatibleTextRenderingDefault(false);
                Application.Run(new frmMain());
            }
        }
    }
}
