using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Aspose.Cells;
using System.IO;

namespace FDCS
{
    public partial class Work01TransferVouch : Form
    {
        public Work01TransferVouch()
        {
            InitializeComponent();
        }

        private void Work01TransferVouch_Load(object sender, EventArgs e)
        {
            //初始化Mapping
            InitLeMapping();
            InitItemMapping();



        }

        private void tsbClose_Click(object sender, EventArgs e)
        {
            Close();
        }

        private void InitLeMapping()
        {
            var cLePath = Application.StartupPath + @"\Mapping\" + Properties.Settings.Default.LEMapping;
            var cLeDir = Application.StartupPath + @"\Mapping";
            if (!Directory.Exists(cLeDir))
            {
                MessageBox.Show("Mapping Folder not exist!");
                return;
            }
            if (!File.Exists(cLePath))
            {
                MessageBox.Show("LE Mapping Excel File not exist!");
                return;
            }
            var workbook = new Workbook(cLePath);
            var cells = workbook.Worksheets[0].Cells;

            //生成Mapping数据源
            for (var i = 1; i < cells.MaxDataRow + 1; i++)
            {
                var leRow = dsMain.LEMapping.NewLEMappingRow();
                leRow.PL = cells[i, 0].StringValue;
                leRow.LE = cells[i, 1].StringValue;
                leRow.Concantenate = cells[i, 2].StringValue;
                leRow.CC = cells[i, 3].StringValue;
                dsMain.LEMapping.Rows.Add(leRow);
            }

        }

        private void InitItemMapping()
        {


            var cItemPath = Application.StartupPath + @"\Mapping\" + Properties.Settings.Default.ITEMMapping;
            var cItemDir = Application.StartupPath + @"\Mapping";
            if (!Directory.Exists(cItemDir))
            {
                MessageBox.Show("Mapping Folder not exist!");
                return;
            }
            if (!File.Exists(cItemPath))
            {
                MessageBox.Show("Item Mapping Excel File not exist!");
                return;
            }
            //打开数据
            var workbook = new Workbook(cItemPath);
            var cells = workbook.Worksheets[0].Cells;

            //生成Mapping数据源
            for (var i = 1; i < cells.MaxDataRow + 1; i++)
            {
                var itemRow = dsMain.ITEMMapping.NewITEMMappingRow();
                itemRow.Item = cells[i, 0].StringValue;
                itemRow.CrAccount = cells[i, 1].StringValue;
                itemRow.CrGETDFolder = cells[i, 2].StringValue;
                itemRow.CrGEIOFolder = cells[i, 3].StringValue;
                itemRow.CrBV = cells[i, 4].StringValue;
                itemRow.CrMktSegment = cells[i, 5].StringValue;
                itemRow.CrDestination = cells[i, 6].StringValue;
                itemRow.CrSource = cells[i, 7].StringValue;
                itemRow.DrAccount = cells[i, 8].StringValue;
                itemRow.DrFolder = cells[i, 9].StringValue;
                itemRow.DrBV = cells[i, 10].StringValue;
                itemRow.DrMktSegment = cells[i, 11].StringValue;
                itemRow.DrDestination = cells[i, 12].StringValue;
                itemRow.DrSource = cells[i, 13].StringValue;
                dsMain.ITEMMapping.Rows.Add(itemRow);
            }
        }

        private void tsbLoadSource_Click(object sender, EventArgs e)
        {
            
            //导入数据
            if (ofdMain.ShowDialog() != DialogResult.OK)
                return;
            if (string.IsNullOrEmpty(ofdMain.FileName))
                return;
            dsMain.DataInput.Rows.Clear();
            var wBook = new Workbook(ofdMain.FileName);
            var cells = wBook.Worksheets[0].Cells;
            pbMain.Value = 0;
            pbMain.Maximum = cells.MaxDataRow;
            tslblProgress.Text = @"Loading File";
            for (var i = 1; i < cells.MaxDataRow + 1; i++)
            {
                var inputRow = dsMain.DataInput.NewDataInputRow();
                inputRow.RowNo=cells[i, 0].StringValue;
                inputRow.OrderNumber = cells[i, 1].StringValue;
                inputRow.LE = cells[i, 2].StringValue;
                inputRow.PL = cells[i, 3].StringValue;
                inputRow.Sub_Mod = cells[i, 4].StringValue;
                inputRow.MOD_Code = cells[i, 5].StringValue;
                inputRow.INS = cells[i, 6].StringValue;
                inputRow.WAR = cells[i, 7].StringValue;
                inputRow.APPI = cells[i, 8].StringValue;
                inputRow.APPII = cells[i, 9].StringValue;
                inputRow.VCP = cells[i, 10].StringValue;
                dsMain.DataInput.Rows.Add(inputRow);
                pbMain.Value = i;
            }

            //校验数据，并转换成2位小数的数值
            pbMain.Value = 0;
            tslblProgress.Text = @"Check Data Format";
            for (var i = 0; i < dsMain.DataInput.Rows.Count ; i++)
            {
                decimal dins,dwar,dappi,dappii,dvcp;
                var cins = dsMain.DataInput.Rows[i]["INS"].ToString();
                var cwar = dsMain.DataInput.Rows[i]["WAR"].ToString();
                var cappi = dsMain.DataInput.Rows[i]["APPI"].ToString();
                var cappii = dsMain.DataInput.Rows[i]["APPII"].ToString();
                var cvcp = dsMain.DataInput.Rows[i]["VCP"].ToString();
                //判断INS
                var bins = decimal.TryParse(cins, out dins);
                if (bins)
                {
                    dsMain.DataInput.Rows[i]["INS"] = Math.Round(dins, 2);
                }
                else
                {
                    uGridInput.Rows[i].Cells["INS"].Appearance.BackColor = Color.Pink;
                }

                //判断WAR
                var bwar = decimal.TryParse(cwar, out dwar);
                if (bwar)
                {
                    dsMain.DataInput.Rows[i]["WAR"] = Math.Round(dwar, 2);
                }
                else
                {
                    uGridInput.Rows[i].Cells["WAR"].Appearance.BackColor = Color.Pink;
                }

                //判断APPI
                var bappi = decimal.TryParse(cappi, out dappi);
                if (bappi)
                {
                    dsMain.DataInput.Rows[i]["APPI"] = Math.Round(dappi, 2);
                }
                else
                {
                    uGridInput.Rows[i].Cells["APPI"].Appearance.BackColor = Color.Pink;
                }


                //判断APPII
                var bappii = decimal.TryParse(cappii, out dappii);
                if (bappii)
                {
                    dsMain.DataInput.Rows[i]["APPII"] = Math.Round(dappii, 2);
                }
                else
                {
                    uGridInput.Rows[i].Cells["APPII"].Appearance.BackColor = Color.Pink;
                }


                //判断INS
                var bvcp = decimal.TryParse(cvcp, out dvcp);
                if (bvcp)
                {
                    dsMain.DataInput.Rows[i]["VCP"] = Math.Round(dvcp, 2);
                }
                else
                {
                    uGridInput.Rows[i].Cells["VCP"].Appearance.BackColor = Color.Pink;
                }

                if (bins & bwar & bappi & bappii & bvcp)
                {
                    uGridInput.Rows[i].Appearance.BackColor = Color.LightBlue;
                }

            }

            tslblProgress.Text = @"Load Box Excel Data Source，Success";
        }

        private void tsbtnTransferData_Click(object sender, EventArgs e)
        {
            for (var i = 0; i < uGridInput.Rows.Count; i++)
            {
                //先写借
                WriteINSDr(i);
                WriteINSCr(i);
            }
        }
        /// <summary>
        /// 输出写借方
        /// </summary>
        /// <param name="i"></param>
        private void WriteINSDr(int i)
        {
            var outputRow = dsMain.DataOutPut.NewDataOutPutRow();
            outputRow.ACCOUNTING_DATE = i.ToString();
            outputRow.CURRENCY_CODE = "";
            outputRow.DATE_CREATED = "";
            outputRow.CURRENCY_CONV_DATE = "";
            outputRow.CURRENCY_CONV_TYPE = "";
            outputRow.CURRENCY_CONV_RATE = "";
            outputRow.AFF_COMPANY = "";
            outputRow.AFF_ACCOUNT = "";
            outputRow.AFF_CENTER = "";
            outputRow.AFF_BASE_VAR = "";
            outputRow.AFF_MODALITY = "";
            outputRow.AFF_MKT_SEGMENT = "";
            outputRow.AFF_FOLDER = "";
            outputRow.AFF_SOURCE = "";
            outputRow.AFF_DESTINATION = "";
            outputRow.ENTERED_DR = "";
            outputRow.ENTERED_CR = "";
            outputRow.ACCOUNTED_DR = "";
            outputRow.ACCOUNTED_CR = "";
            outputRow.SET_OF_BOOKS_ID = "";
            outputRow.ACTUAL_FLAG = "";
            outputRow.AFF_JOURNAL_CATEGORY = "";
            outputRow.AFF_JOURNAL_SOURCE = "";
            outputRow.PERIOD = "";
            outputRow.CODE_COMBINATION_ID = "";
            outputRow.COMPANY_CODE_MAP = "";
            outputRow.JOURNAL_SOURCE_MAP = "";
            outputRow.JOURNAL_CATEGORY_MAP = "";
            outputRow.JOURNAL_BATCH = "";
            outputRow.BATCH_DESCRIPTION = "";
            outputRow.JOURNAL_NAME = "";
            outputRow.JOURNAL_DESCRIPTION = "";
            outputRow.JOURNAL_REFERENCE = "";
            outputRow.JOURNALLINEDESC = "";
            outputRow.LEGACY_ACCOUNT = "";
            outputRow.LEGACY_JRNL_NUM = "";
            outputRow.LEGACY_OFFSET_ACCT = "";
            outputRow.BILL_TO_CUSTOMER = "";
            outputRow.SHIP_TO_CUSTOMER = "";
            outputRow.EMPLOYEE_NUM = "";
            outputRow.INVENTORY_ORG = "";
            outputRow.MA_CODE = "";
            outputRow.MATERIAL_CLASS = "";
            outputRow.PO_ITEM = "";
            outputRow.ORDER_NUM = "";
            outputRow.INV_ITEM_NUM = "";
            outputRow.QUANTITY = "";
            outputRow.UNIT_OF_MEASURE = "";
            outputRow.DOCUMENT_NUM = "";
            outputRow.DOCUMENT_DATE = "";
            outputRow.PROJECT_NUM = "";
            outputRow.DOCUMENT_NUM2 = "";
            outputRow.SHIPPED_DATE = "";
            outputRow.VAT_CODE = "";
            outputRow.ACTUAL_HOURS = "";
            outputRow.CONSIGNMT_CONTRACT = "";
            outputRow.COST_KEY = "";
            outputRow.PO_NUM = "";
            outputRow.PSI_CODE = "";
            outputRow.RETURN_MAT_CODE = "";
            outputRow.TRANSACTION_CODE = "";
            outputRow.VENDOR_NUM = "";
            outputRow.SERVICE_ACCTG_KEY = "";
            outputRow.REFERENCE_AMOUNT = "";
            outputRow.LOCAL_MAPPING_FIELD1 = "";
            outputRow.LOCAL_MAPPING_FIELD2 = "";
            outputRow.LOCAL_MAPPING_FIELD3 = "";
            outputRow.LOCAL_MAPPING_FIELD4 = "";
            outputRow.LOCAL_MAPPING_FIELD5 = "";
            outputRow.LOCAL_MAPPING_FIELD6 = "";
            outputRow.LOCAL_MAPPING_FIELD7 = "";
            outputRow.LOCAL_MAPPING_FIELD8 = "";
            outputRow.LOCAL_MAPPING_FIELD9 = "";
            outputRow.LOCAL_MAPPING_FIELD10 = "";
            outputRow.LOCAL_MAPPING_FIELD11 = "";
            outputRow.LOCAL_MAPPING_FIELD12 = "";
            outputRow.LOCAL_MAPPING_FIELD13 = "";
            outputRow.LOCAL_MAPPING_FIELD14 = "";
            outputRow.LOCAL_MAPPING_FIELD15 = "";
            outputRow.LOCAL_MAPPING_FIELD16 = "";
            outputRow.LOCAL_MAPPING_FIELD17 = "";
            outputRow.LOCAL_MAPPING_FIELD18 = "";
            outputRow.LOCAL_MAPPING_FIELD19 = "";
            dsMain.DataOutPut.Rows.Add(outputRow);
        }

        /// <summary>
        /// 输出写代方
        /// </summary>
        /// <param name="i"></param>
        private void WriteINSCr(int i)
        {
            var outputRow = dsMain.DataOutPut.NewDataOutPutRow();
            outputRow.ACCOUNTING_DATE = "";
            outputRow.CURRENCY_CODE = "";
            outputRow.DATE_CREATED = "";
            outputRow.CURRENCY_CONV_DATE = "";
            outputRow.CURRENCY_CONV_TYPE = "";
            outputRow.CURRENCY_CONV_RATE = "";
            outputRow.AFF_COMPANY = "";
            outputRow.AFF_ACCOUNT = "";
            outputRow.AFF_CENTER = "";
            outputRow.AFF_BASE_VAR = "";
            outputRow.AFF_MODALITY = "";
            outputRow.AFF_MKT_SEGMENT = "";
            outputRow.AFF_FOLDER = "";
            outputRow.AFF_SOURCE = "";
            outputRow.AFF_DESTINATION = "";
            outputRow.ENTERED_DR = "";
            outputRow.ENTERED_CR = "";
            outputRow.ACCOUNTED_DR = "";
            outputRow.ACCOUNTED_CR = "";
            outputRow.SET_OF_BOOKS_ID = "";
            outputRow.ACTUAL_FLAG = "";
            outputRow.AFF_JOURNAL_CATEGORY = "";
            outputRow.AFF_JOURNAL_SOURCE = "";
            outputRow.PERIOD = "";
            outputRow.CODE_COMBINATION_ID = "";
            outputRow.COMPANY_CODE_MAP = "";
            outputRow.JOURNAL_SOURCE_MAP = "";
            outputRow.JOURNAL_CATEGORY_MAP = "";
            outputRow.JOURNAL_BATCH = "";
            outputRow.BATCH_DESCRIPTION = "";
            outputRow.JOURNAL_NAME = "";
            outputRow.JOURNAL_DESCRIPTION = "";
            outputRow.JOURNAL_REFERENCE = "";
            outputRow.JOURNALLINEDESC = "";
            outputRow.LEGACY_ACCOUNT = "";
            outputRow.LEGACY_JRNL_NUM = "";
            outputRow.LEGACY_OFFSET_ACCT = "";
            outputRow.BILL_TO_CUSTOMER = "";
            outputRow.SHIP_TO_CUSTOMER = "";
            outputRow.EMPLOYEE_NUM = "";
            outputRow.INVENTORY_ORG = "";
            outputRow.MA_CODE = "";
            outputRow.MATERIAL_CLASS = "";
            outputRow.PO_ITEM = "";
            outputRow.ORDER_NUM = "";
            outputRow.INV_ITEM_NUM = "";
            outputRow.QUANTITY = "";
            outputRow.UNIT_OF_MEASURE = "";
            outputRow.DOCUMENT_NUM = "";
            outputRow.DOCUMENT_DATE = "";
            outputRow.PROJECT_NUM = "";
            outputRow.DOCUMENT_NUM2 = "";
            outputRow.SHIPPED_DATE = "";
            outputRow.VAT_CODE = "";
            outputRow.ACTUAL_HOURS = "";
            outputRow.CONSIGNMT_CONTRACT = "";
            outputRow.COST_KEY = "";
            outputRow.PO_NUM = "";
            outputRow.PSI_CODE = "";
            outputRow.RETURN_MAT_CODE = "";
            outputRow.TRANSACTION_CODE = "";
            outputRow.VENDOR_NUM = "";
            outputRow.SERVICE_ACCTG_KEY = "";
            outputRow.REFERENCE_AMOUNT = "";
            outputRow.LOCAL_MAPPING_FIELD1 = "";
            outputRow.LOCAL_MAPPING_FIELD2 = "";
            outputRow.LOCAL_MAPPING_FIELD3 = "";
            outputRow.LOCAL_MAPPING_FIELD4 = "";
            outputRow.LOCAL_MAPPING_FIELD5 = "";
            outputRow.LOCAL_MAPPING_FIELD6 = "";
            outputRow.LOCAL_MAPPING_FIELD7 = "";
            outputRow.LOCAL_MAPPING_FIELD8 = "";
            outputRow.LOCAL_MAPPING_FIELD9 = "";
            outputRow.LOCAL_MAPPING_FIELD10 = "";
            outputRow.LOCAL_MAPPING_FIELD11 = "";
            outputRow.LOCAL_MAPPING_FIELD12 = "";
            outputRow.LOCAL_MAPPING_FIELD13 = "";
            outputRow.LOCAL_MAPPING_FIELD14 = "";
            outputRow.LOCAL_MAPPING_FIELD15 = "";
            outputRow.LOCAL_MAPPING_FIELD16 = "";
            outputRow.LOCAL_MAPPING_FIELD17 = "";
            outputRow.LOCAL_MAPPING_FIELD18 = "";
            outputRow.LOCAL_MAPPING_FIELD19 = "";
            dsMain.DataOutPut.Rows.Add(outputRow);
        }
    }
}
