using System;
using System.IO;
using System.Text.RegularExpressions;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using TestMateAppGroup;
using WinFormExcel;
using ExcelProcess = Microsoft.Office.Interop.Excel;
using System.Windows.Forms;

namespace ElecLineStatistics
{
    class ProductionDataConversion
    {

        public string strPanelStorage { private set; get; }
        public string strHaierStorage { private set; get; }
        public string strFinishGoodsStorage { private set; get; }
        public string strPanelConsumptionFolder { private set; get; }
        public string strProductionConfirmFolder { private set; get; }
       
        private PCBType xlsPCBType = new PCBType(ProgramInit.strPCBType_g);

        public ProductionDataConversion()
        {
            strPanelStorage = @ProgramInit.oAppCfg.AppSettings.Settings["PanelStorage"].Value;
            strHaierStorage = @ProgramInit.oAppCfg.AppSettings.Settings["HaierStorage"].Value;
            strFinishGoodsStorage = @ProgramInit.oAppCfg.AppSettings.Settings["FinishGoodsStorage"].Value;

            strPanelConsumptionFolder = @ProgramInit.oAppCfg.AppSettings.Settings["PanelConsumption"].Value;
            strProductionConfirmFolder = @ProgramInit.oAppCfg.AppSettings.Settings["ProductConfirm"].Value;
        }

        public void InitMainUI(TextBox tbxPCBType, TextBox tbxPanelConsumption, TextBox tbxConfirm, TextBox tbxMonth)
        {
            tbxPCBType.Text = xlsPCBType.PCBTypeFile;
            tbxPanelConsumption.Text = strPanelConsumptionFolder;
            tbxConfirm.Text = strProductionConfirmFolder;
            tbxMonth.Text = DateTime.Now.ToString("MM");
        }

        public void GeneratePCBTypeList(string strMonth, string strPanelConsumptionFile, string strConfirmFile)
        {

            //Get Panel Variant Information for Panel Consumption File
            List<string> lstSAPVariant = GetVeriantWithSlocation(strPanelConsumptionFile);

            //Get the Unique Panel Variant List
            string[] ArrPanelStorage = strPanelStorage.Split(';');
            string[] ArrHaierStorage = strHaierStorage.Split(';');

            List<string> lstVariant = FilterSAPVariant(lstSAPVariant, ArrPanelStorage, "Panel");
            //Get Variant List for Haier Shop-In-Shop
            lstVariant.AddRange(FilterSAPVariant(lstSAPVariant, ArrHaierStorage, "HaierShop"));

            lstVariant = lstVariant.Distinct().ToList();
            
            //Get Variant with Confirm StorageLocation
            List<string> lstConfirmVariant = GetVeriantWithSlocation(strConfirmFile);
            string[] ArrFGStorage = strFinishGoodsStorage.Split(';');
            List<string> lstFinishGoodsVaraint = FilterSAPVariant(lstConfirmVariant, ArrFGStorage, "FinishGoods");

            ExcelProcess.Workbook wbPCBType = MsExcelFile.xlsOpenWorkbook(xlsPCBType.PCBTypeFile, true, false, "", ProgramInit.PCBPASSWORD);
            ExcelProcess.Worksheet wsPCBTypeList = wbPCBType.Worksheets[PCBType.TYPESHEET];

            TestMateApp.ProcShow("导入型号信息 ...");
            //Export Panel Information
            UpdateVariantProperties(lstVariant, wsPCBTypeList, "面板需求");
            //Export FinishGoods Information
            UpdateVariantProperties(lstFinishGoodsVaraint, wsPCBTypeList, "电子成品");

            TestMateApp.ProcShow("保存PCBType表 ...");
            //Get the PCB Type Folder
            string PCBFolder = Path.GetDirectoryName(xlsPCBType.xlsPCBType.FilePath);
            string strPCBListFile = Path.GetDirectoryName(PCBFolder) + $"\\TypeList\\PCBTypList_{strMonth}.xlsx";

            //Save the PCB Type List to Another File
            wbPCBType.Save();
            wbPCBType.SaveAs(strPCBListFile);
            wbPCBType.Worksheets["PCBAType"].Delete();

            MsExcelFile.Showup(wbPCBType);
        }

        List<string> FilterSAPVariant(List<string> SAPVariantList, string[] ArrStorageLocation,string strVariantStamp)
        {
            List<string> lstFilterVariant = new List<string>();
            for(int iList = 0; iList < SAPVariantList.Count; iList++)
            {
                string[] ArrEach = SAPVariantList[iList].Split('-');
                if (Array.IndexOf(ArrStorageLocation, ArrEach[1]) >= 0) lstFilterVariant.Add(ArrEach[0] + "-" + strVariantStamp);
            }

            return lstFilterVariant;
        }

        private static string[] GetSAPFileContent(string strSAPFile)
        {
            //Open Consumption File and Get Complete Content
            FileStream fsSAPFile = new FileStream(strSAPFile, FileMode.Open);
            StreamReader srSAPFile = new StreamReader(fsSAPFile, Encoding.Default);

            string strConsumption = srSAPFile.ReadToEnd();

            srSAPFile.Close();
            fsSAPFile.Close();

            //Split File Content into Array
            string[] ArrContent = strConsumption.Split('\r');

            return ArrContent;
        }

        public static List<string> GetVeriantWithSlocation(string PanelConsumptionFile)
        {
            //Split File Content into Array
            string[] ArrContent = GetSAPFileContent(PanelConsumptionFile);
            string[] ArrConsumption = new string[ArrContent.Length];

            int iMaterial = -1, iSloc = -1, iStdQty = -1;
            for (int iLine = 0; iLine < ArrContent.Length; iLine++)
            {
                string[] ArrLine = ArrContent[iLine].Trim().Split('\t');
                //Ge the Material and Storage Location Column Index
                if (ArrLine.Length > 0 && ArrContent[iLine].Contains("Material"))
                {
                    iStdQty = ArrLine.Length;
                    iMaterial = Array.IndexOf(ArrLine, "Material");
                    iSloc = Array.IndexOf(ArrLine, "SLoc");
                }
                else
                {   //Get the Variant and Storage Location
                    if (iStdQty > 0 && ArrLine.Length > Math.Max(iSloc, iMaterial)) ArrConsumption[iLine] = ArrLine[iMaterial] + "-" + ArrLine[iSloc];
                }
            }
            //Filter the Variant and Storage Location Information into Unique List, Remove Empty Items
            List<string> lstVariant = ArrConsumption.ToList().Distinct().ToList();
            lstVariant = (from x in lstVariant where x != null select x).ToList();

            return lstVariant;
        }

        public void UpdateVariantProperties(List<string> lstVariant, ExcelProcess.Worksheet wsPCBTypeList, string strColumnTitle)
        {
            //Get WorkSheet Basic Information
            object[,] ArrWsPCB = wsPCBTypeList.UsedRange.Value2;
            string[] ArrPCBHead = MsExcelFile.GetSheetRowColumn(ArrWsPCB, true, 1, MsExcelFile.xlGetRowColumn.xlGetRow);
            string[] ArrPCBType = MsExcelFile.GetSheetRowColumn(ArrWsPCB, true, 1, MsExcelFile.xlGetRowColumn.xlGetColumn);
            //Get the Storage Column Index
            int iStorage = Array.IndexOf(ArrPCBHead, strColumnTitle) + 1;
            //Load Storage Loction
            object[] ArrStorageLocation = MsExcelFile.GetSheetRowColumn(ArrWsPCB, true, iStorage, MsExcelFile.xlGetRowColumn.xlGetColumn);
            //Merge new Storage Location from Consumption File
            for (int iPCB = 0; iPCB < lstVariant.Count; iPCB++)
            {
                int iPCBIndex = Array.IndexOf(ArrPCBType, lstVariant[iPCB].Split('-')[0]);

                if (iPCBIndex > 0)
                {
                    ArrStorageLocation[iPCBIndex] = lstVariant[iPCB].Split('-')[1];
                }
            }
            //Load Information into WorkSheet
            wsPCBTypeList.Range[wsPCBTypeList.Cells[1, iStorage], wsPCBTypeList.Cells[ArrStorageLocation.Length, iStorage]].Value2 = MsExcelFile.BuildRangeValue(ArrStorageLocation);
        }

    }

    class QDReportConversion
    {

        public QDConvertConfig QDReportConfig { get; private set; }
        public string strCurrentOrgFolder { private set; get; }
        public string strCurrentQDFolder { private set; get; }
        public string strCurrentTime { private set; get; }

        public string strQDConvertConfigFile { private set; get; }

        private object[,] ArrBasicData = null;

        public QDReportConversion(string _strConfigFile)
        {
            strQDConvertConfigFile = _strConfigFile;
        }

        public void InitQDConvertConfig(TextBox tbxQDConfigFile, TextBox tbxOrgSAPFolder, TextBox tbxConvertTime, RadioButton rdbWeek, RadioButton rdbMonth)
        {
            QDReportConfig = new QDConvertConfig(strQDConvertConfigFile);

            tbxQDConfigFile.Text = QDReportConfig.strConvertConfigFile;
            tbxOrgSAPFolder.Text = QDReportConfig.strSAPDataFolder;

            tbxConvertTime.Text = DateTime.Now.ToString("yyyy") + "-CW" + TestMateApp.GetWeekNum(DateTime.Now.AddDays(-7)).ToString().PadLeft(2, '0');

            rdbWeek.Click += SetTimeFrame;
            rdbMonth.Click += SetTimeFrame;
            
            void SetTimeFrame(object sender, EventArgs e)
            {
                RadioButton rdbAct = sender as RadioButton;

                if (rdbAct.Name.Contains("Week"))
                {
                     tbxConvertTime.Text = DateTime.Now.ToString("yyyy") + "-CW" + TestMateApp.GetWeekNum(DateTime.Now.AddDays(-7)).ToString().PadLeft(2, '0');
                }
                else
                {
                    tbxConvertTime.Text = DateTime.Now.AddMonths(-1).ToString("yyyy-MM");
                }
            }
        }

        public void InitQDConvertConfig()
        {
            QDReportConfig = new QDConvertConfig(strQDConvertConfigFile);
        }

        public string SetCurrentFolder(string _strTimeFrame, ListView lvwQDFile)
        {
            strCurrentTime = _strTimeFrame;
            strCurrentOrgFolder = QDReportConfig.strSAPDataFolder + "\\" + _strTimeFrame;
            strCurrentQDFolder = QDReportConfig.strQDInputFolder + _strTimeFrame;

            string strReturnMessage = "";
            if (!Directory.Exists(strCurrentOrgFolder))
            {
                strReturnMessage = "未发现原始质量数据, 请检查文件夹!!!";
            }
            else
            {
                lvwQDFile.Items.Clear();
                FileInfo[] ArrOrgFile = TestMateApp.LoadFileList(strCurrentOrgFolder, "*.xl*");

                //Load the QD SAP File Information 
                foreach(QDConvertConfig.QDProcConverter ProcConverter in QDReportConfig.ArrQDConfigSet)
                {
                    var oFindFile = from x in ArrOrgFile where x.Name.Contains(ProcConverter.ProcName) orderby x.LastAccessTime select x;

                    if (oFindFile.Count() > 0)
                    {
                        FileInfo QDFile = oFindFile.Last();

                        ListViewItem lviItem = lvwQDFile.Items.Add(QDFile.Name);
                        lviItem.SubItems.Add(ProcConverter.ProcName);
                        lviItem.SubItems.Add(QDFile.CreationTime.ToString("yyyy-MM-dd HH:mm:ss"));
                        lviItem.Checked = true;
                    }
                    else
                    {
                        strReturnMessage += (string.IsNullOrEmpty(strReturnMessage) ? "流程缺失: " : "") + ProcConverter.ProcName;
                    }
                }

                if (!Directory.Exists(strCurrentQDFolder)) Directory.CreateDirectory(strCurrentQDFolder);
            }

            return strReturnMessage;
        }

        public void InitVerifyBasicData(ListViewItem lviRepair)
        {
            string strRepairFile = strCurrentOrgFolder + "\\" + lviRepair.Text;

            ExcelProcess.Workbook wbRepair = MsExcelFile.xlsOpenWorkbook(strRepairFile);
            ExcelProcess.Worksheet wsRepBasic = wbRepair.Worksheets["BasicData"];

            ArrBasicData = wsRepBasic.UsedRange.Value2;

            MsExcelFile.Close(wbRepair);
        }

        public void ConvertRepairInReportMode(ListViewItem lviRepFile, PCBType pcbFamily)
        {
            string strRepairFile = strCurrentOrgFolder + "\\" + lviRepFile.Text;
            string strProcName = lviRepFile.SubItems[1].Text;

            QDConvertConfig.QDProcConverter CurrConverter = QDReportConfig.GetConvertConfig(strProcName);

            ExcelProcess.Workbook wbRepair = MsExcelFile.xlsOpenWorkbook(strRepairFile);
            ExcelProcess.Worksheet wsRepair = wbRepair.Worksheets["RepRec"];

            ////Get the Repair Record From the Database
            List<string[]> lstRepRec = MsExcelFile.GetFilterList(MsExcelFile.GetSheetArray(wsRepair));
            lstRepRec.RemoveAt(0);

            //Convert the Repair Record to the Standard Record Output
            List<string[]> lstConvertRep = new List<string[]>();
            List<string> lstPCBQty = new List<string>();

            lstConvertRep.Add(CurrConverter.ArrSAPHead);
            lstPCBQty.Add("PCB-Qty");

            foreach (string[] ArrRepLine in lstRepRec)
            {
                //Get the Fixed Item From the Repair Record
                string[] ArrTemp = new string[CurrConverter.ArrSAPHead.Length];
                Array.Copy(ArrRepLine, ArrTemp, 8);
                //Build the Record From Each Failure Item
                for (int iFailureCol = 8; iFailureCol < 24; iFailureCol += 4)
                {
                    if (!string.IsNullOrEmpty(ArrRepLine[iFailureCol]))
                    {
                        Array.Copy(ArrRepLine, iFailureCol, ArrTemp, 8, 4);
                        lstConvertRep.Add(ArrTemp);

                        //Add the PCB Quantity for the First Failure
                        lstPCBQty.Add(iFailureCol == 8 ? "1" : "0");

                    }
                    else
                    {
                        break;
                    }
                }
            }
            //Convert the Repair Report to Quality Input Structure
            List<string[]> lstQDInput = GetConvertRepSheet(lstConvertRep, CurrConverter);

            //Convert Empty Column and the Quanity Column for the Repair Data Output
            int iEmptyCol = TestMateApp.GetArrayIndex(CurrConverter.ArrProcHead, "Comp.prod");
            int iFailureQtyCol = TestMateApp.GetArrayIndex(CurrConverter.ArrProcHead, "F-qty");
            int iPCBQtyCol = TestMateApp.GetArrayIndex(CurrConverter.ArrProcHead, "PCB-qty");

            for (int iRow = 1; iRow < lstQDInput.Count; iRow++)
            {
                lstQDInput[iRow][iEmptyCol] = "";
                lstQDInput[iRow][iFailureQtyCol] = "1";
                lstQDInput[iRow][iPCBQtyCol] = lstPCBQty[iRow];
            }

            //Get ICT Column Data (SMD is Include on Repair Basic Data)
            string[] ArrICTSMD = MsExcelFile.GetColumn(ArrBasicData, "ICT", true, true);
            //Get ICT, FIT, SMD Array List
            string[] ArrICT = (from x in ArrICTSMD where !x.ToUpper().Contains("SMD") select x).ToArray();
            string[] ArrSMD = (from x in ArrICTSMD where x.ToUpper().Contains("SMD") select x).ToArray();
            string[] ArrFIT = MsExcelFile.GetColumn(ArrBasicData, "FIT", true, true);
            //Merge All Line Array
            string[] ArrMachine = ArrICTSMD.Concat(ArrFIT).ToArray();

            //Change All Line Data to Machine Type
            int iLineCol = TestMateApp.GetArrayIndex(CurrConverter.ArrProcHead, "Line");
            for (int iRow = 1; iRow < lstQDInput.Count; iRow++)
            {
                string strLineMachine = lstQDInput[iRow][iLineCol];
                string strMacType = GetMachineType(strLineMachine);
                lstQDInput[iRow][iLineCol] = strMacType;

                //Get Machine Type
                string GetMachineType(string strMachineName)
                {
                    int iMachineIndex = TestMateApp.GetArrayIndex(ArrMachine, strMachineName);

                    string strType = "WrongMachine";
                    if (iMachineIndex >= 0)
                    {
                        if (iMachineIndex < ArrICT.Length)
                        {
                            strType = "ICT";
                        }
                        else if (iMachineIndex < (ArrSMD.Length + ArrICT.Length))
                        {
                            strType = "SMD";
                        }
                        else if (iMachineIndex >= (ArrSMD.Length + ArrICT.Length))
                        {
                            strType = "FIT";
                        }
                    }

                    return strType;
                }
            }

            //Filter the Wrong Machine Items
            lstQDInput = (from x in lstQDInput where x[iLineCol] != "WrongMachine" select x).ToList();

            //Verify QDReport with Process/Family Name ...
            lstQDInput = VerifyQDReport(lstQDInput, CurrConverter, pcbFamily);
            SaveQDReportFile(lstQDInput, "RepSMD");


            //Remove the Repair Date Column from the Repair QD File
            int iRepDateCol = TestMateApp.GetArrayIndex(CurrConverter.ArrProcHead, "Repair date");

            List<string[]> lstRepair = new List<string[]>();
            foreach (string[] ArrQDInput in lstQDInput)
            {
                if (ArrQDInput[iLineCol] != "SMD")
                {
                    List<string> lstQDItem = ArrQDInput.ToList();

                    lstQDItem.RemoveAt(iRepDateCol);
                    lstRepair.Add(lstQDItem.ToArray());
                }
            }

            SaveQDReportFile(lstRepair, CurrConverter.ProcName);

            MsExcelFile.Close(wbRepair);

        }

        public void ConverQDFile(ListViewItem lviQDFile, PCBType pcbFamily)
        {
            string strOrgQDFile = strCurrentOrgFolder + "\\" + lviQDFile.Text;
            string strProcName = lviQDFile.SubItems[1].Text;

            QDConvertConfig.QDProcConverter CurrConverter = QDReportConfig.GetConvertConfig(strProcName);

            ExcelProcess.Workbook wbSAP = MsExcelFile.xlsOpenWorkbook(strOrgQDFile, false);
            ExcelProcess.Worksheet wsSAP = wbSAP.ActiveSheet;

            List<string[]> lstQDInput = new List<string[]>();
            if (wsSAP.UsedRange.Count > CurrConverter.ArrSAPHead.Length)
            {
                object[,] ArrWSOrg = wsSAP.UsedRange.Value2;
                List<string[]> lstOrgItem = MsExcelFile.GetFilterList(ArrWSOrg);

                //Convert Report to QD Input Items
                lstQDInput = GetConvertRepSheet(lstOrgItem, CurrConverter);
                lstQDInput = VerifyQDReport(lstQDInput, CurrConverter, pcbFamily);
            }
            else
            {
                lstQDInput.Add(CurrConverter.ArrProcHead);
            }

            SaveQDReportFile(lstQDInput, CurrConverter.ProcName);
            MsExcelFile.Close(wbSAP);
        }

        public void ConvertRepairSAPUploadFile(string strSMDRepairFile)
        {

            //Get Repair Converter Settings for Reapair Process
            QDConvertConfig.QDProcConverter RepConfig = QDReportConfig.GetConvertConfig("Repair");
            List<int> lstSAPIndex = GetRepairSAPUploadIndex(RepConfig);

            //Open the Repair SMD File and Load Repair Item
            ExcelProcess.Workbook wbRepSMD = MsExcelFile.xlsOpenWorkbook(strSMDRepairFile);
            ExcelProcess.Worksheet wsRepSMD = wbRepSMD.ActiveSheet;
            List<string[]> lstRepairRecord = MsExcelFile.GetFilterList(MsExcelFile.GetSheetArray(wsRepSMD));
            //Close SMD Repair File
            MsExcelFile.Close(wbRepSMD);

            //Start Convert on the SMD Repair Record
            lstRepairRecord.RemoveAt(0);
            List<string[]> lstSAPUploadItem = new List<string[]>();
            foreach(string[] ArrRepairItem in lstRepairRecord)
            {
                string[] ArrItem = new string[RepConfig.ArrUploadHead.Length];
                for(int iHead = 0; iHead < lstSAPIndex.Count; iHead++)
                {

                    int iIndex = lstSAPIndex[iHead];

                    switch (iIndex)
                    {

                        case -2:  //Constant Column

                            ArrItem[iHead] = RepConfig.ArrUpLoadSequence[iHead].Replace(">", "").Replace("<", "");
                            break;

                        case -1:  //Null Column
                            ArrItem[iHead] = string.Empty;
                            break;

                        default: //Normal Repair Item

                            ArrItem[iHead] = ArrRepairItem[lstSAPIndex[iHead]];
                            break;
                    }
                }
                lstSAPUploadItem.Add(ArrItem);
            }

            //Finalize the Record List
            lstSAPUploadItem = FinalizeSAPUploadRecord(lstSAPUploadItem, RepConfig);
            string strTimeFolder = Path.GetFileName(Path.GetDirectoryName(strSMDRepairFile));
            string strRepairUpLoadFile = QDReportConfig.strSAPUploadFolder + $"RepairSAPUpload_{strTimeFolder}.txt";

            //Add the Upload Head
            lstSAPUploadItem.Insert(0, RepConfig.ArrUploadHead);

            //Save Repair Upload File
            if (File.Exists(strRepairUpLoadFile)) File.Delete(strRepairUpLoadFile);
            TestMateApp.SaveListArray2TempFile(lstSAPUploadItem, strRepairUpLoadFile, true, "\t");

        }


        private List<string[]> FinalizeSAPUploadRecord(List<string[]> SAPUploadList, QDConvertConfig.QDProcConverter RepConfig)
        {
            int iFailureCode = TestMateApp.GetArrayIndex(RepConfig.ArrUpLoadSequence, "F-Mode");
            string[] ArrFilterCode = TMConfiguration.INIFile.INIGetStringValue(QDReportConfig.strConvertConfigFile, "Repair", "IgnoreFailureCode", "").Split(',');

            //Filter Item with the Ignored Failure Code
            List<string[]> lstSAPRecord = (from x in SAPUploadList where !ArrFilterCode.Any(t => x[iFailureCode].StartsWith(t)) select x).ToList();

            //Change the Failure Code Like 1127s
            lstSAPRecord = (from x in lstSAPRecord select UpdateFailureCode(x, iFailureCode)).ToList();

            //Get the DateColumn from the Record
            List<int> lstDateListIndex = new List<int>();
            for(int iHead = 0; iHead < RepConfig.ArrUpLoadSequence.Length; iHead++)
            {
                if (RepConfig.ArrUpLoadSequence[iHead].ToUpper().Contains("DATE"))
                {
                    lstDateListIndex.Add(iHead);
                }
            }

            //Convert the Repair Date and Test Date Information
            lstSAPRecord = (from x in lstSAPRecord select ConvertDateSection(x)).ToList();

            //Convert the Test Ident
            int iTestIdentCol = TestMateApp.GetArrayIndex(RepConfig.ArrUploadHead, "Test ident");
            lstSAPRecord = (from x in lstSAPRecord select ConvertTestIdent(x, iTestIdentCol)).ToList();

            return lstSAPRecord;

            string[] ConvertDateSection(string[] ArrItem)
            {
                foreach(int iHead in lstDateListIndex)
                {
                    if(double.TryParse(ArrItem[iHead],out double dblDate))
                    {
                        DateTime dtConvertDate = DateTime.FromOADate(dblDate);
                        ArrItem[iHead] = dtConvertDate.ToString("dd.MM.yyyy");
                    }
                }
                return ArrItem;
            }

            string[] ConvertTestIdent(string[] ArrItem,int iColumn)
            {
                switch (ArrItem[iColumn])
                {
                    case "SMD":
                        ArrItem[iColumn] = "REF";
                        break;
                    default:

                        ArrItem[iColumn] = "IN" + ArrItem[iColumn].Substring(0, 1);
                        break;
                }

                return ArrItem;
            }

            string[] UpdateFailureCode(string[] ArrItem, int iCodeColumn)
            {
                if (ArrItem[iFailureCode].Length > 4)
                {
                    ArrItem[iFailureCode] = ArrItem[iFailureCode].Substring(0, 4);
                }
                else if(ArrItem[iFailureCode].Length == 3)
                {
                    ArrItem[iFailureCode] = ArrItem[iFailureCode].PadLeft(4, '0');
                }

               return ArrItem;
            }
        }

        private List<int> GetRepairSAPUploadIndex(QDConvertConfig.QDProcConverter RepConfig)
        {

            string[] ArrSAPHead = RepConfig.ArrUpLoadSequence;
            string[] ArrSMDRepairHead = RepConfig.ArrProcHead;

            List<int> lstSAPHeadIndex = new List<int>();
            foreach(string strUploadHead in ArrSAPHead)
            {
                if (strUploadHead.ToUpper().Equals("{NULL]"))
                {
                    lstSAPHeadIndex.Add(-1);
                }
                else if (strUploadHead.StartsWith("<") && strUploadHead.EndsWith(">")) 
                {
                    lstSAPHeadIndex.Add(-2);
                }
                else
                {
                    int iHead = TestMateApp.GetArrayIndex(ArrSMDRepairHead, strUploadHead);
                    lstSAPHeadIndex.Add(iHead);
                }
            }

            return lstSAPHeadIndex;

        }

        private void SaveQDReportFile(List<string[]> lstQDInput, string strProcName)
        {

            //Create New WorkBook
            ExcelProcess.Application oXlApp = new ExcelProcess.Application();
            ExcelProcess.Workbook wbQDInput = oXlApp.Workbooks.Add();
            ExcelProcess.Worksheet wsQDInput = wbQDInput.ActiveSheet;
            //Export Information to QD Data Sheet
            MsExcelFile.ExportRangeValue(wsQDInput, lstQDInput, MsExcelFile.xlImportMode.xlOverwrite);

            wsQDInput.EnableAutoFilter = true;
            wsQDInput.Columns.AutoFit();

            //Get the File Name With Version
            string strFilePath = TestMateApp.GetNewVersionFile(strCurrentQDFolder + "\\" + strProcName + " " + strCurrentTime + ".xlsx");
            string strOldFile = TestMateApp.GetLatestFile(strCurrentQDFolder, $"{strProcName}*.xl*");

            //Delete the Old Process File
            if (!string.IsNullOrEmpty(strOldFile)) File.Delete(strOldFile);
            //Save the Converted File
            wbQDInput.SaveAs(strFilePath);

            MsExcelFile.Close(wbQDInput, true);
        }

        private List<string[]> VerifyQDReport(List<string[]> lstQDInput, QDConvertConfig.QDProcConverter CurrConverter, PCBType pcbFamily)
        {
            //Add Process Head
            lstQDInput[0] = CurrConverter.ArrProcHead;

            //Add the Family Name
            lstQDInput = GenerateRepFamilyName(lstQDInput, pcbFamily);
            lstQDInput = GenerateRepProcess(lstQDInput, pcbFamily);

            if (CurrConverter.ProcName.ToUpper().Equals("REPAIR"))
            {
                int iLineCol = TestMateApp.GetArrayIndex(CurrConverter.ArrProcHead, "Line");

                for (int iQDLine = 1; iQDLine < lstQDInput.Count; iQDLine++)
                {
                    string[] ArrQDLine = lstQDInput[iQDLine];
                    string strProcess = ArrQDLine[ArrQDLine.Length - 1];

                    string strLine = ArrQDLine[iLineCol];

                    //If the Process Name Contains in the Processes List
                    if (Array.IndexOf(CurrConverter.ArrProcess, strProcess) >= 0)
                    {
                        //If the Process Is SMD, Check if the Line is "SMD-xx"
                        if (strLine.ToUpper().Contains("SMD"))
                        {
                            if (strProcess.ToUpper().Contains("SMD")) lstQDInput[iQDLine][ArrQDLine.Length - 1] = "";
                        }
                        else
                        {
                            lstQDInput[iQDLine][ArrQDLine.Length - 1] = "";
                        }
                    }
                }
            }
            else
            {
                for (int iQDLine = 1; iQDLine < lstQDInput.Count; iQDLine++)
                {
                    string[] ArrQDLine = lstQDInput[iQDLine];
                    string strProcess = ArrQDLine[ArrQDLine.Length - 1];
                    if (Array.IndexOf(CurrConverter.ArrProcess, strProcess) >= 0) lstQDInput[iQDLine][ArrQDLine.Length - 1] = "";
                }
            }

            lstQDInput = CheckFailureCode(lstQDInput);
            return lstQDInput;
        }

        private List<string[]> GetConvertRepSheet(List<string[]> lstOrgItem, QDConvertConfig.QDProcConverter CurrConverter)
        {
            //Filter the Empty Rows
            var vNonEmpty = from x in lstOrgItem where x[0] != null select x;
            lstOrgItem = vNonEmpty.ToList();
            //Get Head of SAP File
            string[] ArrSAPHead = lstOrgItem[0];

            //Get Sequence Index From SAP Head Array
            int[] ArrConvIndex = new int[CurrConverter.ArrSequence.Length];

            List<int> lstDateCol = new List<int>();
            for (int iSeq = 0; iSeq < ArrConvIndex.Length; iSeq++)
            {
                string strSeqHead = CurrConverter.ArrSequence[iSeq];
                ArrConvIndex[iSeq] = Array.IndexOf(ArrSAPHead, strSeqHead);

                //Get Head of Date Column
                if (strSeqHead.ToUpper().Contains("DATE") | strSeqHead.ToUpper().Contains("DAY") | strSeqHead.ToUpper().Contains("日期")) lstDateCol.Add(iSeq);

            }

            //Adjust the Date Format for Excel Sheet
            List<string[]> lstQDInput = new List<string[]>();
            for (int iRow = 0; iRow < lstOrgItem.Count; iRow++)
            {
                string[] ArrOrgLine = lstOrgItem[iRow];
                string[] ArrQDLine = new string[CurrConverter.ArrProcHead.Length];

                for (int iIndex = 0; iIndex < ArrConvIndex.Length; iIndex++)
                {
                    ArrQDLine[iIndex] = ArrOrgLine[ArrConvIndex[iIndex]];
                }

                //Convert Date Value of Excel to Date String Format
                foreach (int iDateCol in lstDateCol)
                {
                    if (double.TryParse(ArrQDLine[iDateCol], out double dblOADate))
                    {
                        ArrQDLine[iDateCol] = DateTime.FromOADate(dblOADate).ToString("yyyy-MM-dd");
                    }
                    else
                    {
                        string[] ArrDate = ArrQDLine[iDateCol].Replace('.', '-').Replace('_', '-').Replace(' ', '-').Split('-');

                        //Rebuild the Date String with Sections in the String
                        if (ArrDate.Length == 3)
                        {
                            if (ArrDate[0].Length == 4)
                            {
                                ArrQDLine[iDateCol] = ArrDate[0] + "-" + ArrDate[1] + "-" + ArrDate[2];
                            }
                            else
                            {
                                ArrQDLine[iDateCol] = ArrDate[2] + "-" + ArrDate[1] + "-" + ArrDate[0];
                            }
                        }
                    }
                }

                //Cleanup the "#" in Records
                for (int iCol = 0; iCol < ArrQDLine.Length; iCol++)
                {
                    if (ArrQDLine[iCol] == "#" | ArrQDLine[iCol] == "0") ArrQDLine[iCol] = "";
                }

                lstQDInput.Add((string[])ArrQDLine.Clone());
            }

            return lstQDInput;

        }

        private List<string[]> GenerateRepFamilyName(List<string[]> lstQDInput, PCBType pcbFamily)
        {
            string[] ArrRepHead = lstQDInput[0];

            int iMaterialCol = TestMateApp.GetArrayIndex(ArrRepHead, "Material");
            int iFamilyCol = TestMateApp.GetArrayIndex(ArrRepHead, "Family");

            //Get Family Name for Each Record
            List<string[]> lstQDFamily = new List<string[]>();
            foreach (string[] ArrInput in lstQDInput)
            {
                if (string.IsNullOrEmpty(ArrInput[iFamilyCol]))
                {
                    ArrInput[iFamilyCol] = pcbFamily.GetFamilyName(ArrInput[iMaterialCol]);
                }

                lstQDFamily.Add((string[])ArrInput.Clone());
            }

            return lstQDFamily;
        }

        private List<string[]> GenerateRepProcess(List<string[]> lstQDInput, PCBType pcbFamily)
        {
            string[] ArrRepHead = lstQDInput[0];
            int iMaterialCol = TestMateApp.GetArrayIndex(ArrRepHead, "Material");

            //Get Family Name for Each Record
            List<string[]> lstProcess = new List<string[]>();
            for (int iRow = 0; iRow < lstQDInput.Count; iRow++)
            {
                string[] ArrInput = lstQDInput[iRow];

                string[] ArrTemp = new string[ArrInput.Length + 1];
                ArrInput.CopyTo(ArrTemp, 0);
                if (iRow != 0)
                {
                    ArrTemp[ArrTemp.Length - 1] = pcbFamily.GetProcessName(ArrInput[iMaterialCol]);
                }
                else
                {
                    ArrTemp[ArrTemp.Length - 1] = "Process";
                }

                lstProcess.Add((string[])ArrTemp.Clone());
            }

            return lstProcess;
        }

        private List<string[]> CheckFailureCode(List<string[]> lstQDInput)
        {
            //Get the Failure Mode Column
            int iFailureCodeCol = TestMateApp.GetArrayIndex(lstQDInput[0], "F-mode");
            string[] ArrFailureCode = MsExcelFile.GetColumn(ArrBasicData, "缺陷代码");

            List<string[]> lstFailureCode = new List<string[]>();
            for (int iRow = 0; iRow < lstQDInput.Count; iRow++)
            {
                string[] ArrInput = lstQDInput[iRow];
                string[] ArrTemp = new string[ArrInput.Length + 1];

                ArrInput.CopyTo(ArrTemp, 0);
                //Check the Failure Code from the Failure Code DataBase
                if (iRow != 0)
                {
                    string strFailureCode = ArrInput[iFailureCodeCol];
                    if (!string.IsNullOrEmpty(strFailureCode))
                    {
                        strFailureCode = strFailureCode.PadLeft(4, '0');
                        int iFailureRow = Array.IndexOf(ArrFailureCode, strFailureCode);
                        if (iFailureRow < 0) ArrTemp[ArrTemp.Length - 1] = "Wrong";
                    }

                }
                else
                {
                    ArrTemp[ArrTemp.Length - 1] = "Failure Code";
                }

                lstFailureCode.Add(ArrTemp);
            }

            return lstFailureCode;
        }

        public class QDConvertConfig
        {
            public string strConvertConfigFile { private set; get; }
            public string strSAPDataFolder { private set; get; }
            public string strQDInputFolder { private set; get; }
            public string strSAPUploadFolder { private set; get; }

            public QDProcConverter[] ArrQDConfigSet { private set; get; }

            public struct QDProcConverter
            {
                public string ProcName;
                public string[] ArrSAPHead;
                public string[] ArrProcHead;
                public string[] ArrSequence;
                public string[] ArrProcess;

                public string[] ArrUploadHead;
                public string[] ArrUpLoadSequence;
            }

            public QDConvertConfig(string _strConfigFile)
            {

                //Get Config File Path, SAP Report Head and Quality Source Data Head List
                strConvertConfigFile = _strConfigFile;
                strSAPDataFolder = TMConfiguration.INIFile.INIGetStringValue(_strConfigFile, "QualityStatistics", "SAPOrgFolder", "");
                strQDInputFolder = TMConfiguration.INIFile.INIGetStringValue(_strConfigFile, "QualityStatistics", "QDStatFolder", "");
                strSAPUploadFolder = TMConfiguration.INIFile.INIGetStringValue(_strConfigFile, "QualityStatistics", "SAPUploadFolder", "");

                string[] ArrProcName = TMConfiguration.INIFile.INIGetStringValue(_strConfigFile, "QualityStatistics", "Process", "").Split(',');

                //Initial Process Config Settings from Config File
                ArrQDConfigSet = new QDProcConverter[ArrProcName.Length];
                //Get Each Process Configuration from File
                for (int iProc = 0; iProc < ArrProcName.Length; iProc++)
                {

                    string strProcName = ArrProcName[iProc];

                    string[] ArrProcHead = TMConfiguration.INIFile.INIGetStringValue(_strConfigFile, strProcName, "ProcHead", "").Split(',');
                    string[] ArrSAPHead = TMConfiguration.INIFile.INIGetStringValue(_strConfigFile, strProcName, "SAPHead", "").Split(',');
                    string[] ArrSequence = TMConfiguration.INIFile.INIGetStringValue(_strConfigFile, strProcName, "Sequence", "").Split(',');
                    string[] ArrProcess = TMConfiguration.INIFile.INIGetStringValue(_strConfigFile, strProcName, "PCBProcess", "").Split(',');

                    string[] ArrUploadHead = TMConfiguration.INIFile.INIGetStringValue(_strConfigFile, strProcName, "SAPUploadHead", "").Split(',');
                    string[] ArrUploadSequence = TMConfiguration.INIFile.INIGetStringValue(_strConfigFile, strProcName, "UploadSequence", "").Split(',');

                    ArrQDConfigSet[iProc].ProcName = strProcName;
                    ArrQDConfigSet[iProc].ArrProcHead = (string[])ArrProcHead.Clone();
                    ArrQDConfigSet[iProc].ArrSAPHead = (string[])ArrSAPHead.Clone();
                    ArrQDConfigSet[iProc].ArrSequence = (string[])ArrSequence.Clone();
                    ArrQDConfigSet[iProc].ArrProcess = (string[])ArrProcess.Clone();

                    ArrQDConfigSet[iProc].ArrUploadHead = (string[])ArrUploadHead.Clone();
                    ArrQDConfigSet[iProc].ArrUpLoadSequence = (string[])ArrUploadSequence.Clone();
                }
            }

            public QDProcConverter GetConvertConfig(string strProcName)
            {
                var oQDConfig = (from x in ArrQDConfigSet where x.ProcName == strProcName select x).First();

                return oQDConfig;
            }

            public string GetConfigValue(string strSection, string strKey)
            {
                return TMConfiguration.INIFile.INIGetStringValue(strConvertConfigFile, strSection, strKey, "");
            }
        }

    }

}
