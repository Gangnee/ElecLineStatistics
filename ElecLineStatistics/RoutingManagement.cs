using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Text.RegularExpressions;
using TestMateAppGroup;
using System.Windows.Forms;
using WinFormExcel;
using ExcelProc = Microsoft.Office.Interop.Excel;

namespace ElecLineStatistics
{
    class RoutingManagement
    {
        /// <summary>
        /// Statistics Database File Path
        /// </summary>
        public string strStatDataBase { get { return ProgramInit.strStatRoutingDataBase_g; } }

        /// <summary>
        /// Standard Routing Database File Path
        /// </summary>
        public string strStandardRoutingDataBase { get { return ProgramInit.strStandardRoutingDataBase_g; } }

        /// <summary>
        /// Process Config FIle Folder
        /// </summary>
        public string strProcConfigFolder { get { return Application.StartupPath + "\\ProcConfig"; } }

        public const string STDROUTINGSHEET = "ProductionRoutingHour";

        /// <summary>
        /// Flag for Routing Initialized
        /// </summary>
        public bool IsRoutingInitialized { get { return StatRoutingList != null; } }

        /// <summary>
        /// Config File List from Each Process
        /// </summary>
        public string[] ArrConfigFile { private set; get; }

        /// <summary>
        /// Process Name in Config File (Chinese)
        /// </summary>
        public string[] ArrProcName { private set; get; }

        /// <summary>
        /// Config Setting from Select Process
        /// </summary>
        public RoutingConfig CurrentRoutingSet = null;

        /// <summary>
        /// Database Head Information of the Standard Routing Database
        /// </summary>
        public string[] ArrStandardDBHead
        {
            get
            {
                string[] ArrHead = { "产品型号", "产品族名", "流程类别", "单件工时", "成本中心", "工时标识", "每组人数", "人机比例", "最大组数", "小时过板数", "报表工时", "工时备注", "拼板数量", "全拼工时", "元件数量", "工时测量", "难度编号", "资源编号", "工时报表", "操作时间" };
                return ArrHead;
            }
        }

        /// <summary>
        /// Routing Matrix From Database for Efficiency Database
        /// </summary>
        public List<string[]> StatRoutingList { private set; get; }

        /// <summary>
        /// Routing List from Standard Routing DataBase
        /// </summary>
        public List<string[]> StandardRoutingList { private set; get; }


        public RoutingManagement()
        {
            FileInfo[] ArrCFGFile = TestMateApp.LoadFileList(strProcConfigFolder, "*.ini");
            //Get Config File List and Process Name List
            ArrConfigFile = (from x in ArrCFGFile select x.FullName).ToArray();
            ArrProcName = (from x in ArrCFGFile select TMConfiguration.INIFile.INIGetStringValue(x.FullName, "General", "ProcName", "ERROR")).ToArray();
        }

        /// <summary>
        /// InitMain UI At Startup of the Program
        /// </summary>
        /// <param name="cmbProcList"></param>
        /// <param name="dgvStdRouting"></param>
        public void InitMainUI(ComboBox cmbProcList, DataGridView dgvStdRouting)
        {
            cmbProcList.Items.AddRange(ArrProcName);

            //Init DataGridView Columns of Standard Routing Database
            TestMateApp.InitDataGridView(ArrStandardDBHead, dgvStdRouting);
        }

        /// <summary>
        /// Init Routing Management for Process Loading
        /// </summary>
        /// <param name="strProcName"></param>
        /// <param name="pnlInput"></param>
        /// <param name="dgvStatDatabase"></param>
        /// <param name="chkMultiLine"></param>
        /// <param name="dgvStatRouting"></param>
        public void InitRoutingProcess(string strProcName, Panel pnlInput, DataGridView dgvStatDatabase, CheckBox chkMultiLine, DataGridView dgvStatRouting)
        {

            //Clear All the UI Information
            TestMateApp.ResetInputBox(pnlInput);
            dgvStatDatabase.Columns.Clear();
            dgvStatRouting.Rows.Clear();

            TestMateApp.ProcShow("初始化流程配置 ...");
            //Set Input Flag for all Input Boxes
            TestMateApp.SetInputBoxFlag(pnlInput);
            //Get Process Configuration for the Selected Routing Process
            string strConfigFile = ArrConfigFile[Array.IndexOf(ArrProcName, strProcName)];

            //Init Routing Process Configuration
            CurrentRoutingSet = new RoutingConfig(strConfigFile);

            //Get the Routing Matrix From Statistic Routing DataBase
            TestMateApp.ProcShow("导入效率工时 ...");
            //Get Statistic Routing DataBase
            StatRoutingList = GetStatRoutingMatrix();

            TestMateApp.ProcShow("导入标准工时 ...");
            StandardRoutingList = GetStandardRougingList();

            TestMateApp.ProcShow("界面初始化 ...");
            //Init Combo Input List
            InitComboInput(pnlInput);

            TestMateApp.SetInputSequence(pnlInput);

            //Enable Flag for MultiLine Input
            chkMultiLine.Checked = false;
            chkMultiLine.Enabled = CurrentRoutingSet.IsMultiLineInput;

            //Init DataGridView for Statistic Routing Data
            TestMateApp.InitDataGridView(StatRoutingList[0], dgvStatDatabase);
        }

        /// <summary>
        /// Query Exist Routing from Efficiency DataBase
        /// </summary>
        /// <param name="strQueryKey"></param>
        /// <param name="dgvStandardRouting"></param>
        /// <param name="dgvStatRouting"></param>
        public void QueryRouting(string strQueryKey, DataGridView dgvStandardRouting, DataGridView dgvStatRouting)
        {
            //Get the Routing Head from Statistic Database File
            string[] ArrRoutingHead = StatRoutingList[0];
            dgvStatRouting.Rows.Clear();

            int iMaterialCol = Array.IndexOf(ArrRoutingHead, "物料号");
            int iFamilyCol = Array.IndexOf(ArrRoutingHead, "族名");
            //Get the Routing which Type or Family Contains Key Value
            var oRoutingSearch = from x in StatRoutingList where x[iMaterialCol] .Contains(strQueryKey) | x[iFamilyCol].ToLower().Contains(strQueryKey.ToLower()) select x;

            //Add the Stat Routing into the DataGridView
            if (oRoutingSearch.Count() > 0)
            {
                foreach(string[] ArrRouting in oRoutingSearch)
                {
                    dgvStatRouting.Rows.Add(ArrRouting);
                }
            }
            
        }

        /// <summary>
        /// Convert Single Efficiency Routing to Standard Routing List
        /// </summary>
        /// <param name="ArrStatRouting"></param>
        /// <returns></returns>
        public List<string[]> GenerateStandardRoutingList(string[] ArrStatRouting)
        {

            //Get Statistics Routing Database Head
            string[] ArrStatHead = StatRoutingList[0];

            //Get Common Config for the Statistical Routings
            string strRoutingName = ArrStatRouting[0];
            string strFamilyName = ArrStatRouting[1];
            string strDataBaseSheet = CurrentRoutingSet.strStatDataBaseSheet;

            string strVariant = CurrentRoutingSet.GetVariantFromRouting(strRoutingName, CurrentRoutingSet.IsVariantWithRevision);
            string strSuffix = CurrentRoutingSet.GetSuffixFromRouting(strRoutingName, CurrentRoutingSet.IsVariantWithRevision);


            List<string[]> lstStandardRouting = new List<string[]>();
            for (int iRouting = CurrentRoutingSet.iRoutingStartFromStatMatrixColumn; iRouting < ArrStatRouting.Length; iRouting++)
            {

                //Get the Routing Hour From Statistics DataBase
                string strRoutingHour = ArrStatRouting[iRouting];

                //Skip Empty Routing Hour
                if (string.IsNullOrWhiteSpace(strRoutingHour) || !double.TryParse(strRoutingHour, out double dblRouting)) continue;

                string[] ArrDBRouting = new string[ArrStandardDBHead.Length];

                ArrDBRouting[GetDBHeadColumn("产品型号")] = strVariant;
                ArrDBRouting[GetDBHeadColumn("产品族名")] = strFamilyName;
                ArrDBRouting[GetDBHeadColumn("报表工时")] = strRoutingName;
                ArrDBRouting[GetDBHeadColumn("工时报表")] = strDataBaseSheet;

                switch (CurrentRoutingSet.strProcKey)
                {

                    case "SMD":

                        ArrDBRouting[GetDBHeadColumn("全拼工时")] = dblRouting.ToString();
                        ArrDBRouting[GetDBHeadColumn("成本中心")] = CurrentRoutingSet.GetRoutingCostCenter();
                        ArrDBRouting[GetDBHeadColumn("流程类别")] = ArrStatHead[iRouting];
                        ArrDBRouting[GetDBHeadColumn("每组人数")] = CurrentRoutingSet.iFixLineOpNumber.ToString();
                        ArrDBRouting[GetDBHeadColumn("工时标识")] = CurrentRoutingSet.GetRoutingFlag(strSuffix);

                        ArrDBRouting[GetDBHeadColumn("工时测量")] = ArrStatRouting[3];

                        //Try to Convert the Standard Work Hour of SMD
                        try
                        {
                            int iPanel = int.Parse(ArrStatRouting[5]);
                            int iCompQty = int.Parse(ArrStatRouting[6]);

                            ArrDBRouting[GetDBHeadColumn("元件数量")] = iCompQty.ToString();
                            ArrDBRouting[GetDBHeadColumn("拼板数量")] = iPanel.ToString();

                            double dbSMDRouting = dblRouting / 3600 / iPanel;

                            ArrDBRouting[GetDBHeadColumn("单件工时")] = dbSMDRouting.ToString();
                        }
                        catch { continue; }  //Return NUll if the Routing Could not be Converted

                        break;

                    case "MWS":

                        ArrDBRouting[GetDBHeadColumn("单件工时")] = dblRouting.ToString();
                        string strRoutingProcess = CurrentRoutingSet.GetRoutingFlag(strSuffix);
                        ArrDBRouting[GetDBHeadColumn("工时标识")] = strRoutingProcess;
                        ArrDBRouting[GetDBHeadColumn("成本中心")] = CurrentRoutingSet.GetRoutingCostCenter(strRoutingProcess);
                        ArrDBRouting[GetDBHeadColumn("每组人数")] = CurrentRoutingSet.GetRoutingPersonQuantity(strRoutingName);
                        ArrDBRouting[GetDBHeadColumn("流程类别")] = CurrentRoutingSet.GetRoutingFlag(strSuffix);

                        break;

                    case "SMDP":
                    case "TSTCell":

                        ArrDBRouting[GetDBHeadColumn("单件工时")] = dblRouting.ToString();
                        ArrDBRouting[GetDBHeadColumn("流程类别")] = ArrStatHead[iRouting];
                        ArrDBRouting[GetDBHeadColumn("成本中心")] = CurrentRoutingSet.GetRoutingCostCenter(ArrStatHead[iRouting]);
                        ArrDBRouting[GetDBHeadColumn("工时标识")] = CurrentRoutingSet.GetRoutingFlag(strSuffix);

                        break;

                    case "SWS":
                    case "POT":
                    case "VAR":
                    case "VCD":

                        ArrDBRouting[GetDBHeadColumn("单件工时")] = dblRouting.ToString();
                        ArrDBRouting[GetDBHeadColumn("流程类别")] = ArrStatHead[iRouting];
                        ArrDBRouting[GetDBHeadColumn("成本中心")] = CurrentRoutingSet.GetRoutingCostCenter();
                        ArrDBRouting[GetDBHeadColumn("工时标识")] = CurrentRoutingSet.GetRoutingFlag(strSuffix);

                        if (CurrentRoutingSet.iFixLineOpNumber < 0)
                        {
                            ArrDBRouting[GetDBHeadColumn("每组人数")] = CurrentRoutingSet.GetRoutingPersonQuantity(strRoutingName);
                        }
                        else
                        {
                            ArrDBRouting[GetDBHeadColumn("每组人数")] = CurrentRoutingSet.iFixLineOpNumber.ToString();
                        }

                        break;

                }

                lstStandardRouting.Add((string[])ArrDBRouting.Clone());

            }



            return lstStandardRouting;
        }

        /// <summary>
        /// UpLoad Routing to Input Box for Editing
        /// </summary>
        /// <param name="dgvStandard"></param>
        /// <param name="pnlInput"></param>
        public void EditRouting(DataGridView dgvStandard, Panel pnlInput)
        {
            if (dgvStandard.SelectedRows.Count == 0) return;

            DataGridViewRow dgrSelect = dgvStandard.SelectedRows[0];
            string[] ArrStdRouting = TestMateApp.GetDataGridViewRow(dgrSelect);

            for(int iRouting = 0; iRouting < ArrStdRouting.Length; iRouting++)
            {
                string strFlag = ArrStandardDBHead[iRouting];

                Control ctlInput = null;
                switch (strFlag)
                {
                    case "单件工时":

                        ctlInput = TestMateApp.GetInputBoxFromPanel(pnlInput, "千件工时");
                        if (ctlInput != null) ctlInput.Text = (double.Parse(ArrStdRouting[iRouting]) * 1000).ToString();

                        break;

                    case "工时标识":

                        ctlInput = TestMateApp.GetInputBoxFromPanel(pnlInput, "工时后缀");
                        if (ctlInput != null) ctlInput.Text = CurrentRoutingSet.GetFlagSuffix(ArrStdRouting[iRouting]);

                        break;

                    default:

                        ctlInput = TestMateApp.GetInputBoxFromPanel(pnlInput, strFlag);
                        if (ctlInput != null) ctlInput.Text = ArrStdRouting[iRouting];

                        break;
                }


                if (strFlag == "单件工时")
                {

                }
                else
                {

                }
            }

        }


        /******************************************************************************************************************************************
         *                                            Private Classs Functions 
         *****************************************************************************************************************************************/

        private List<string[]> GetStatRoutingMatrix()
        {

            ExcelProc.Workbook wbStatDatabase = MsExcelFile.xlsOpenWorkbook(strStatDataBase);
            ExcelProc.Worksheet wsProc = wbStatDatabase.Worksheets[CurrentRoutingSet.strStatDataBaseSheet];

            object[,] ArrStatDatabase = wsProc.UsedRange.Value2;
            string[] ArrStatRTHead = MsExcelFile.GetSheetRowColumn(ArrStatDatabase, true, 2);

            int iMaterialCol = Array.IndexOf(ArrStatRTHead, "物料号");
            //Get the Routing Area Index
            List<int> lstIndex = new List<int>();
            for (int iHeadCol = iMaterialCol; iHeadCol < ArrStatRTHead.Length; iHeadCol++) 
            {
                if (!string.IsNullOrWhiteSpace(ArrStatRTHead[iHeadCol])) lstIndex.Add(iHeadCol + 1);
            }

            //Get Routing DataBase for Effiency Input
            List<string[]> StatRoutingList = MsExcelFile.GetFilterList(ArrStatDatabase, lstIndex.ToArray());

            //Remove First Row From Database
            StatRoutingList.RemoveAt(0);

            //Rebuild the Date From Double to Standard Date String
            if (CurrentRoutingSet.strProcKey.Equals("SMD"))
            {
                int iCommentCol = Array.IndexOf(StatRoutingList[0], "备注");
                int iDateCol = Array.IndexOf(StatRoutingList[0], "更新日期");

                StatRoutingList = (from x in StatRoutingList select BuildDataStringFromRouting(x, iCommentCol)).ToList();
                StatRoutingList = (from x in StatRoutingList select BuildDataStringFromRouting(x, iDateCol)).ToList();
            }


            MsExcelFile.Close(wbStatDatabase);
            return StatRoutingList;

            //Sub Function to Modify Double Date Value to Date String
            string[] BuildDataStringFromRouting(string[] ArrRouting, int iDateColum)
            {
                if(double.TryParse(ArrRouting[iDateColum],out double dblDate))
                {
                    string strDate = DateTime.FromOADate(dblDate).ToString("yyyy-MM-dd");
                    ArrRouting[iDateColum] = strDate;
                }

                return ArrRouting;
            }
        }

        private List<string[]> GetStandardRougingList(bool IsReload=false)
        {
            //If Routing Loaded Return Exist Standard List
            if (StandardRoutingList != null && !IsReload) return StandardRoutingList;

            ExcelProc.Workbook wbStandardRouting = MsExcelFile.xlsOpenWorkbook(strStandardRoutingDataBase);
            ExcelProc.Worksheet wsStandardRouting = wbStandardRouting.Worksheets[STDROUTINGSHEET];

            List<string[]> lstRoutingList = MsExcelFile.GetFilterList(MsExcelFile.GetSheetArray(wsStandardRouting));

            MsExcelFile.Close(wbStandardRouting);

            return lstRoutingList;
        }

        private void InitComboInput(Panel pnlInput)
        {

            bool IsSMDPartVisible = CurrentRoutingSet.strProcKey.Equals("SMD");

            TestMateApp.GetComboBoxFromPanel(pnlInput, "工时测量").Enabled = IsSMDPartVisible;
            TestMateApp.GetTextBoxFromPanel(pnlInput, "拼板数量").Enabled = IsSMDPartVisible;
            TestMateApp.GetTextBoxFromPanel(pnlInput, "全拼工时").Enabled = IsSMDPartVisible;
            TestMateApp.GetTextBoxFromPanel(pnlInput, "元件数量").Enabled = IsSMDPartVisible;


            //Init All Combo Item List
            AddComboList(TestMateApp.GetComboBoxFromPanel(pnlInput, "成本中心"), CurrentRoutingSet.ArrCostCenter);
            AddComboList(TestMateApp.GetComboBoxFromPanel(pnlInput, "工时后缀"), CurrentRoutingSet.ArrRoutingSuffix);
            AddComboList(TestMateApp.GetComboBoxFromPanel(pnlInput, "工时测量"), CurrentRoutingSet.ArrRoutingMeasure);

            //Get 
            if (CurrentRoutingSet.strSubMode.ToUpper().Equals("LINE"))
            {
                AddComboList(TestMateApp.GetComboBoxFromPanel(pnlInput, "流程类别"), TestMateApp.SubArray(StatRoutingList[0], CurrentRoutingSet.iRoutingStartFromStatMatrixColumn));
            }
            else
            {
                AddComboList(TestMateApp.GetComboBoxFromPanel(pnlInput, "流程类别"), CurrentRoutingSet.ArrSuffixFlag);
            }


            void AddComboList(ComboBox cmbSection, string[] ArrList)
            {
                cmbSection.Items.Clear();
                cmbSection.Items.AddRange(ArrList);
            }
        }

        private int GetDBHeadColumn(string strHeadName)
        {
            return Array.IndexOf(ArrStandardDBHead, strHeadName);
        }

    }

    public class RoutingConfig
    {

        private struct CostCenterSet
        {
            public string[] ArrProcess;
            public string strCostCenter;
        }


        public string strRoutingConfigFile { private set; get; }

        public string strProcName { private set; get; }
        public string strProcKey { private set; get; }
        public string strRoutingSheet { private set; get; }
        public string strStatDataBaseSheet { private set; get; }

        public string[] ArrCostCenter { private set; get; }
        public string strSubMode { private set; get; }

        /// <summary>
        /// Routing Data Start Column from the Statistic Routing DataBase
        /// </summary>
        public int iRoutingStartFromStatMatrixColumn { private set; get; }

        /// <summary>
        /// Operator Nr for Each Production Line
        /// </summary>
        public int iFixLineOpNumber { private set; get; }

        public bool IsVariantWithRevision { private set; get; }
        public string[] ArrRoutingSuffix { private set; get; }
        public string[] ArrSuffixFlag { private set;get;}
        public string[] ArrRoutingMeasure { private set; get; }

        public bool IsMultiLineInput { private set; get; }

        private CostCenterSet[] ArrCostCentConfig;

        //Init Routing Settings from Selected Configuration File
        public RoutingConfig(string strConfigFile)
        {
            strRoutingConfigFile = strConfigFile;
            InitRoutingConfig(strConfigFile);
        }

        private void InitRoutingConfig(string strConfigFile)
        {

            //Get Process Name and SheetName in Routing DataBase
            strProcName = GetProcessConfig(strConfigFile, "General", "ProcName");
            strProcKey = GetProcessConfig(strConfigFile, "General", "ProcKey");

            strRoutingSheet = GetProcessConfig(strConfigFile, "General", "SheetName");

            //Get Statisitics Database Configuration
            strStatDataBaseSheet = GetProcessConfig(strConfigFile, "StatDataBase", "DatabaseSheet");
            strSubMode = GetProcessConfig(strConfigFile, "General", "RoutingMode");

            //Get Standard Routing Database Configuration
            ArrCostCenter = GetProcessConfig(strConfigFile, "StandardDataBase", "CostCenter", true);

            iRoutingStartFromStatMatrixColumn = GetConfigNr("StatDataBase", "RoutingStartFromMatrixColumn");
            iFixLineOpNumber = GetConfigNr("StandardDataBase", "FixLineOperator");

            string strWithRevision = GetProcessConfig(strConfigFile, "StandardDataBase", "WithRevision");
            IsVariantWithRevision = strWithRevision.Equals("1");

            //Get Routing Name Configuration
            ArrRoutingSuffix = GetProcessConfig(strConfigFile, "Routing", "Suffix", true);
            ArrSuffixFlag = GetProcessConfig(strConfigFile, "Routing", "Equipment", true);
            ArrRoutingMeasure = GetProcessConfig(strConfigFile, "Routing", "Measure", true);
            IsMultiLineInput = GetProcessConfig(strConfigFile, "Routing", "MultiLine").Equals("1");

            //Get CostCenter Settings
            ArrCostCentConfig = new RoutingConfig.CostCenterSet[ArrCostCenter.Length];
            for (int iCostCenter = 0; iCostCenter < ArrCostCenter.Length; iCostCenter++) 
            {
                string[] ArrProcess = GetProcessConfig(strConfigFile, "CostCenter", ArrCostCenter[iCostCenter], true);

                ArrCostCentConfig[iCostCenter].strCostCenter = ArrCostCenter[iCostCenter];
                ArrCostCentConfig[iCostCenter].ArrProcess = (string[])ArrProcess.Clone();
            }

            int GetConfigNr(string strSection, string strKey)
            {
                if (int.TryParse(GetProcessConfig(strConfigFile, strSection, strKey), out int iOutNr))
                {
                    return iOutNr;
                }
                else
                {
                    return -1;
                }
            }
        }

        private string GetProcessConfig(string strConfigFile, string strSectionName, string strKeyName)
        {
            return TMConfiguration.INIFile.INIGetStringValue(strConfigFile, strSectionName, strKeyName, "ERROR");
        }

        private string[] GetProcessConfig(string strConfigFile, string strSectionName, string strKeyName, bool IsArray)
        {
            return TMConfiguration.INIFile.INIGetStringValue(strConfigFile, strSectionName, strKeyName, "ERROR").Split(',');
        }

        /// <summary>
        /// Get Routing Flag Comment from Suffix
        /// </summary>
        /// <param name="strSuffix"></param>
        /// <returns></returns>
        public string GetRoutingFlag(string strSuffix)
        {
            int iSuffix = Array.IndexOf(ArrRoutingSuffix, strSuffix);

            if (iSuffix >= 0) 
            { 
                return ArrSuffixFlag[iSuffix]; 
            }
            else
            {
                return string.IsNullOrEmpty(strSuffix) ? "" : "N/A";
            }

        }

        /// <summary>
        /// Get Routing Flag From Routing Suffix
        /// </summary>
        /// <param name="strRoutingFlag"></param>
        /// <returns></returns>
        public string GetFlagSuffix(string strRoutingFlag)
        {
            int iSuffix = Array.IndexOf(ArrSuffixFlag, strRoutingFlag);

            if (iSuffix >= 0)
            {
                return ArrRoutingSuffix[iSuffix];
            }
            else
            {
                return "";
            }

        }

        /// <summary>
        /// Get Cost Center for Current Process, If Process Empty Return Fixed CostCenter
        /// </summary>
        /// <param name="strProcess"></param>
        /// <returns></returns>
        public string GetRoutingCostCenter(string strProcess = "")
        {
            if (string.IsNullOrEmpty(strProcess))
            {
                return ArrCostCenter[0];
            }
            else
            {
                var oCostCenter = from x in ArrCostCentConfig where Array.IndexOf(x.ArrProcess, strProcess) >= 0 select x.strCostCenter;
                return oCostCenter.FirstOrDefault();
            }
        }

        /// <summary>
        /// Get Person Quantity From Routing(e.g. MWS Process
        /// </summary>
        /// <param name="strRoutingName"></param>
        /// <returns></returns>
        public string GetRoutingPersonQuantity(string strRoutingName)
        {
            string[] ArrRouting = strRoutingName.Split('-');

            string strPersonQty = "";
            if (ArrRouting.Length == 3) strPersonQty = ArrRouting[2];
            return strPersonQty;
        }

        /// <summary>
        /// Get Suffix from the Routing Name of Efficiency Database
        /// </summary>
        /// <param name="strRouting"></param>
        /// <param name="IsVariantWithVersion"></param>
        /// <returns></returns>
        public string GetSuffixFromRouting(string strRouting, bool IsVariantWithVersion = false)
        {
            string[] ArrRouting = strRouting.Replace('_', '-').Split('-');

            //If no Suffix Return Empty
            if (ArrRouting.Length == 1) return "";

            string strSuffix;
            if (IsVariantWithVersion)
            {
                strSuffix = ArrRouting.Length > 2 ? ArrRouting[ArrRouting.Length - 1] : "";
            }
            else
            {
                strSuffix = ArrRouting[1];
            }

            return strSuffix;
        }

        /// <summary>
        /// Get Variant Information From Efficiency Routing Name
        /// </summary>
        /// <param name="strRoutingName"></param>
        /// <param name="IsWithRevision"></param>
        /// <returns></returns>
        public string GetVariantFromRouting(string strRoutingName, bool IsWithRevision = false)
        {
            string[] ArrRouting = strRoutingName.Replace('-', '_').Split('_');

            if (IsWithRevision)
            {
                return ArrRouting[0] + "-" + ArrRouting[1];
            }
            else
            {
                return Regex.Match(strRoutingName, "[0-9]{5,6}").Value;
            }
        }

    }
}
