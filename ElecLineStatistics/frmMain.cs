using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using TestMateAppGroup;
using WinFormExcel;

namespace ElecLineStatistics
{
    public partial class frmMain : Form
    {
        //Production Conversion Class
        ProductionDataConversion ProdConverter = new ProductionDataConversion();
        //Quanlity Conversion Class
        QDReportConversion QDConverter = new QDReportConversion(Application.StartupPath + "\\StatConfig\\QualityDataConfig.ini");
        //PCB Type Information Class
        PCBType xlxPCBType = new PCBType(ProgramInit.strPCBType_g);

        RoutingManagement RoutingManager = new RoutingManagement();

        public frmMain()
        {
            InitializeComponent();
        }

        private void ExportPanelVariant(object sender, EventArgs e)
        {
            //Get the Consumption and Config File From SAP System
            string strPanelConsumptionFile = TestMateApp.GetLatestFile(ProdConverter.strPanelConsumptionFolder, $"*Consumption*{tbxTypeMonth.Text}.xl*");
            string strConfirmFile = TestMateApp.GetLatestFile(ProdConverter.strProductionConfirmFolder, $"*Confirm*{tbxTypeMonth.Text}*.xl*");

            if (string.IsNullOrEmpty(strPanelConsumptionFile))
            {
                MessageBox.Show(this, "没有面板消耗文件!!!", "生成面板型号列表", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            TestMateApp.ProcShow("载入SAP相关数据 ...", this);
            //Get PCB List Information from PCB Type List
            ProdConverter.GeneratePCBTypeList(tbxTypeMonth.Text, strPanelConsumptionFile, strConfirmFile);

            TestMateApp.ProcShow();
        }

        private void InitMainUI(object sender, EventArgs e)
        {
            lblVersion.Text = ProgramInit.oAppInformation.UpdateDate + " " + ProgramInit.oAppInformation.appVersion;

            //Init Production Converter UI Element
            ProdConverter.InitMainUI(tbxPCBType, tbxPanelConsumption, tbxSAPConfirm, tbxTypeMonth);
            //Init Quanlity Data from SAP to QD Report Source Data
            QDConverter.InitQDConvertConfig(tbxQDConfigFile, tbxQDSAPFolder, tbxQDTimeFrame, rdbQDWeek, rdbQDMonth);
            //Init Main UI For Routing Management
            RoutingManager.InitMainUI(cmbRoutingGroup, dgvStdRouting);
        }

        private void LoadQDFile(object sender, EventArgs e)
        {
            string strMessage = QDConverter.SetCurrentFolder(tbxQDTimeFrame.Text, lvwQDFile);

            if (!string.IsNullOrEmpty(strMessage))
            {
                MessageBox.Show(this, strMessage, "质量文件转换", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

        }

        private void ConvertQDInputFile(object sender, EventArgs e)
        {

            if (string.IsNullOrEmpty(QDConverter.strCurrentOrgFolder))
            {
                MessageBox.Show(this, "流程没有初始化, 无法转换!!!", "质量文件转换", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            TestMateApp.ProcShow("初始化缺陷代码 ...", this);

            ListViewItem lviRepairFile = null;
            foreach(ListViewItem lviQDFile in lvwQDFile.Items)
            {
                if (lviQDFile.SubItems[1].Text.ToUpper() == "REPAIR") lviRepairFile = lviQDFile;
            }

            if (lviRepairFile == null)
            {
                MessageBox.Show(this, "没有维修文件, 无法转换!!!", "质量文件转换", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            else
            {
                QDConverter.InitVerifyBasicData(lviRepairFile);
            }

            foreach (ListViewItem lviQDFile in lvwQDFile.Items)
            {
                TestMateApp.ProcShow($"转换质量数据 {lviQDFile.SubItems[1].Text} ...");

                if (lviQDFile.SubItems[1].Text == "Repair")
                {
                    QDConverter.ConvertRepairInReportMode(lviQDFile, xlxPCBType);
                }
                else
                {
                    QDConverter.ConverQDFile(lviQDFile, xlxPCBType);
                }
            }

            TestMateApp.ProcShow();

        }

        private void CreateRepairUploadFile(object sender, EventArgs e)
        {
            //Get the Converted SMD File
            string strRepSMDFile = TestMateApp.GetLatestFile(QDConverter.strCurrentQDFolder, "RepSMD*.xl*");

            if (string.IsNullOrEmpty(strRepSMDFile))
            {
                MessageBox.Show(this, "没有贴片维修文件, 无法转换!!!", "质量文件转换", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            TestMateApp.ProcShow("转换贴片维修文件 ...", this);
            //Create the SMD UpLoad File
            QDConverter.ConvertRepairSAPUploadFile(strRepSMDFile);

            TestMateApp.ProcShow();

        }

        private void InitRoutingInput(object sender, EventArgs e)
        {
            TestMateApp.ProcShow("", this);
            tbxRTQueryKey.Text = null;
            RoutingManager.InitRoutingProcess(cmbRoutingGroup.Text, pnlRoutingInput, dgvStatRouting, chkMultiLine, dgvStdRouting);
            TestMateApp.ProcShow();
        }

        private void RoutingInputClear(object sender, EventArgs e)
        {
            TestMateApp.ResetInputBox(pnlRoutingInput, null, true);
        }

        private void QueryRouting(object sender, EventArgs e)
        {
            if (!RoutingManager.IsRoutingInitialized) { MessageBox.Show(this, "工时没有初始化!!!", "工时搜索", MessageBoxButtons.OK, MessageBoxIcon.Error); return; }
            if (string.IsNullOrWhiteSpace(tbxRTQueryKey.Text)) { MessageBox.Show(this, "请输入搜索条件!!!", "工时搜索", MessageBoxButtons.OK, MessageBoxIcon.Error); return; }

            RoutingManager.QueryRouting(tbxRTQueryKey.Text, dgvStdRouting, dgvStatRouting);
        }

        private void ViewStandardRoutingFile(object sender, EventArgs e)
        {
            MsExcelFile.Showup(RoutingManager.strStandardRoutingDataBase, "打开标准工时数据库 ...", true);
        }

        private void ViewStatisticsDataBase(object sender, EventArgs e)
        {
            MsExcelFile.Showup(RoutingManager.strStatDataBase, "打开效率工时数据库 ...", true,ProgramInit.ROUTINGPASSWORD);
        }

        private void QueryRoutingOnKeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar.Equals((char)13)) btnQueryRouting.PerformClick();
        }

        private void ViewStatRouting(object sender, EventArgs e)
        {
            dgvStdRouting.Rows.Clear();
            if (dgvStatRouting.SelectedRows.Count == 0) return;

            //Clear History Loaded Information
            dgvStdRouting.Rows.Clear();
            TestMateApp.ResetInputBox(pnlRoutingInput);

            DataGridViewRow dgrSelect = dgvStatRouting.SelectedRows[0];
            //Buildup Statistical DataBaseRouting
            string[] ArrStatRouting = TestMateApp.GetDataGridViewRow(dgrSelect);
            //Get Standard Routing List
            List<string[]> lstDBRouting = RoutingManager.GenerateStandardRoutingList(ArrStatRouting);

            foreach (string[] ArrDBRouting in lstDBRouting)
            {
                dgvStdRouting.Rows.Add(ArrDBRouting);
            }

        }

        private void ViewStandardRouting(object sender, EventArgs e)
        {
            RoutingManager.EditRouting(dgvStdRouting, pnlRoutingInput);
        }
    }
}
