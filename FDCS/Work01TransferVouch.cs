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
using System.Reflection;
using Infragistics.Documents.Excel;

namespace FDCS
{
    public partial class Work01TransferVouch : Form
    {
        decimal dMor;
        public Work01TransferVouch()
        {
            InitializeComponent();
        }

        private void Work01TransferVouch_Load(object sender, EventArgs e)
        {
            txtDESC_INS.Text = Properties.Settings.Default.DESC_INS;
            txtDESC_WAR.Text = Properties.Settings.Default.DESC_WAR;
            txtDESC_APPI.Text = Properties.Settings.Default.DESC_APPI;
            txtDESC_APPII.Text = Properties.Settings.Default.DESC_APPII;
            txtDESC_VCP.Text = Properties.Settings.Default.DESC_VCP;
            //初始化Mapping
            InitLeMapping();
            InitItemMapping();
            InitSDMatrix();


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
            var workbook = new Aspose.Cells.Workbook(cLePath);
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


        private void InitSDMatrix()
        {
            var cLePath = Application.StartupPath + @"\Mapping\" + Properties.Settings.Default.SDMatrix;
            var cLeDir = Application.StartupPath + @"\Mapping";
            if (!Directory.Exists(cLeDir))
            {
                MessageBox.Show("Mapping Folder not exist!");
                return;
            }
            if (!File.Exists(cLePath))
            {
                MessageBox.Show("SDMatrix Mapping Excel File not exist!");
                return;
            }
            var workbook = new Aspose.Cells.Workbook(cLePath);
            var cells = workbook.Worksheets[0].Cells;

            //生成Mapping数据源
            for (var i = 1; i < cells.MaxDataRow + 1; i++)
            {
                var sdRow = dsMain.SDMatrix.NewSDMatrixRow();
                sdRow.Item = cells[i, 0].StringValue;
                sdRow.PL = cells[i, 1].StringValue;
                sdRow.LE = cells[i, 2].StringValue;
                sdRow.SubMod = cells[i, 3].StringValue;
                sdRow.ModCode = cells[i, 4].StringValue;
                sdRow.DrAccount = cells[i, 5].StringValue;
                sdRow.DrCC = cells[i, 6].StringValue;
                sdRow.DrFolder = cells[i, 7].StringValue;
                sdRow.CrPL = cells[i, 8].StringValue;
                sdRow.CrSubMod = cells[i, 9].StringValue;
                sdRow.CrModCode = cells[i, 10].StringValue;
                sdRow.CrAccount = cells[i, 11].StringValue;
                sdRow.CrCC = cells[i, 12].StringValue;
                sdRow.CrFolder = cells[i, 13].StringValue;

                dsMain.SDMatrix.Rows.Add(sdRow);
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
            var workbook = new Aspose.Cells.Workbook(cItemPath);
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
        /// <summary>
        /// 载入Input数据源
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void tsbLoadSource_Click(object sender, EventArgs e)
        {

            //导入数据
            if (ofdMain.ShowDialog() != DialogResult.OK)
                return;
            if (string.IsNullOrEmpty(ofdMain.FileName))
                return;
            dsMain.DataInput.Rows.Clear();
            var wBook = new Aspose.Cells.Workbook(ofdMain.FileName);
            var cells = wBook.Worksheets[0].Cells;
            pbMain.Value = 0;
            pbMain.Maximum = cells.MaxDataRow;
            tslblProgress.Text = @"Loading File";
            for (var i = 1; i < cells.MaxDataRow + 1; i++)
            {
                var inputRow = dsMain.DataInput.NewDataInputRow();
                inputRow.RowNo = cells[i, 0].StringValue;
                inputRow.OrderNumber = cells[i, 1].StringValue;
                inputRow.LE = cells[i, 2].StringValue;
                inputRow.PL = cells[i, 3].StringValue;
                inputRow.Sub_Mod = cells[i, 4].StringValue;
                inputRow.MOD_Code = cells[i, 5].StringValue;

                decimal dins, dwar, dappi, dappii, dvcp;
                var cins = cells[i, 6].StringValue;
                var cwar = cells[i, 7].StringValue;
                var cappi = cells[i, 8].StringValue;
                var cappii = cells[i, 9].StringValue;
                var cvcp = cells[i, 10].StringValue;
                //判断INS
                var bins = decimal.TryParse(cins, out dins);
                if (bins)
                {
                    //inputRow.INS = cells[i, 6].StringValue;
                    inputRow.INS = Math.Round(dins, 2);
                }
                else
                {
                    inputRow.INS = 0;

                    inputRow.SetColumnError("INS", "SouceData: " + cins + "    cannot convert to decimal");
                }

                //判断WAR
                var bwar = decimal.TryParse(cwar, out dwar);
                if (bwar)
                {
                    inputRow.WAR = Math.Round(dwar, 2);
                }
                else
                {
                    inputRow.WAR = 0;
                    inputRow.SetColumnError("WAR", "SouceData: " + cwar + "   cannot convert to decimal");
                }

                //判断APPI
                var bappi = decimal.TryParse(cappi, out dappi);
                if (bappi)
                {
                    inputRow.APPI = Math.Round(dappi, 2);
                }
                else
                {
                    inputRow.APPI = 0;
                    inputRow.SetColumnError("APPI", "SouceData: " + cappi + "  cannot convert to decimal");
                }


                //判断APPII
                var bappii = decimal.TryParse(cappii, out dappii);
                if (bappii)
                {
                    inputRow.APPII = Math.Round(dappii, 2);
                }
                else
                {
                    inputRow.APPII = 0;
                    inputRow.SetColumnError("APPII", "SouceData: " + cappii + "    cannot convert to decimal");
                }


                //判断INS
                var bvcp = decimal.TryParse(cvcp, out dvcp);
                if (bvcp)
                {
                    inputRow.VCP = Math.Round(dvcp, 2);
                }
                else
                {
                    inputRow.VCP = 0;
                    inputRow.SetColumnError("VCP", "SouceData: " + cvcp + "   cannot convert to decimal");
                }


                dsMain.DataInput.Rows.Add(inputRow);
                pbMain.Value = i;
            }

            //校验数据，并转换成2位小数的数值
            pbMain.Value = 0;
            tslblProgress.Text = @"Check Data Format";
            for (var i = 0; i < dsMain.DataInput.Rows.Count; i++)
            {
                decimal dins, dwar, dappi, dappii, dvcp;
                var cins = dsMain.DataInput.Rows[i]["INS"].ToString();
                var cwar = dsMain.DataInput.Rows[i]["WAR"].ToString();
                var cappi = dsMain.DataInput.Rows[i]["APPI"].ToString();
                var cappii = dsMain.DataInput.Rows[i]["APPII"].ToString();
                var cvcp = dsMain.DataInput.Rows[i]["VCP"].ToString();
                //判断INS
                var bins = cins.Equals("0");
                if (bins)
                {
                    uGridInput.Rows[i].Cells["INS"].Appearance.BackColor = Color.Pink;
                }

                //判断WAR
                var bwar = cwar.Equals("0");
                if (bwar)
                {
                    uGridInput.Rows[i].Cells["WAR"].Appearance.BackColor = Color.Pink;
                }

                //判断APPI
                var bappi = cappi.Equals("0");
                if (bappi)
                {
                    uGridInput.Rows[i].Cells["APPI"].Appearance.BackColor = Color.Pink;
                }


                //判断APPII
                var bappii = cappii.Equals("0");
                if (bappii)
                {
                    uGridInput.Rows[i].Cells["APPII"].Appearance.BackColor = Color.Pink;
                }


                //判断INS
                var bvcp = cvcp.Equals("0");
                if (bvcp)
                {
                    uGridInput.Rows[i].Cells["VCP"].Appearance.BackColor = Color.Pink;
                }

                if (bins & bwar & bappi & bappii & bvcp)
                {
                    uGridInput.Rows[i].Appearance.BackColor = Color.Pink;
                }

            }

            tslblProgress.Text = @"Load Box Excel Data Source，Success";
            if (dsMain.DataInput.HasErrors)
            {
                tslblProgress.Text = @"Input DataSouce Find Some Error, Cannot covert some data to Decimal Type";
            }
        }


        /// <summary>
        /// 点击了Transfer Data按钮
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void tsbtnTransferData_Click(object sender, EventArgs e)
        {
            //判断MOR是否有效
            if (uneMor.Value == null || string.IsNullOrEmpty(uneMor.Value.ToString()))
            {
                MessageBox.Show("Error:MOR is incorrect");
                return;
            }
            if (!decimal.TryParse(uneMor.Value.ToString(), out dMor))
            {
                MessageBox.Show("Error:MOR is incorrect");
                return;
            }
            if (dMor <= 0)
            {
                MessageBox.Show("Error:MOR Must be 0+");
                return;
            }

            dsMain.DataOutPut.Rows.Clear();
            tcMain.SelectedTab = tcMain.TabPages[1];

            //执行五个Item的生成
            var iInsCount = GenerateINS();
            var iWarCount = GenerateWAR();
            var iAPPI = GenerateAPPI();
            var iAPPII = APPIIGenerateAPPII();
            var iVCP = GenerateVCP();


            if (uGridOutput.Rows.Count < 1)
                return;
            uGridOutput.Rows[iInsCount].Appearance.BackColor = Color.LightGreen;
            uGridOutput.Rows[iInsCount].Activate();
            Application.DoEvents();

            uGridOutput.Rows[iInsCount + iWarCount].Appearance.BackColor = Color.Orange;
            uGridOutput.Rows[iInsCount + iWarCount].Activate();
            Application.DoEvents();


            uGridOutput.Rows[iInsCount + iWarCount + iAPPI].Appearance.BackColor = Color.Pink;
            uGridOutput.Rows[iInsCount + iWarCount + iAPPI].Activate();
            Application.DoEvents();

            uGridOutput.Rows[iInsCount + iWarCount + iAPPI + iAPPII].Appearance.BackColor = Color.LightBlue;
            uGridOutput.Rows[iInsCount + iWarCount + iAPPI + iAPPII].Activate();
            Application.DoEvents();

            if (uGridOutput.Rows.Count<=(iInsCount + iWarCount + iAPPI + iAPPII + iVCP))
            {
                uGridOutput.Rows[uGridOutput.Rows.Count-1].Appearance.BackColor = Color.LightYellow;
                uGridOutput.Rows[uGridOutput.Rows.Count - 1].Activate();
            }
            else
            {
                uGridOutput.Rows[iInsCount + iWarCount + iAPPI + iAPPII + iVCP].Appearance.BackColor = Color.LightYellow;
                uGridOutput.Rows[iInsCount + iWarCount + iAPPI + iAPPII + iVCP].Activate();
            }
            
            Application.DoEvents();

            tslblProgress.Text = "Transfer Complete";

        }

        /// <summary>
        /// 生成Installationo数据
        /// </summary>
        /// <returns></returns>
        private int GenerateINS()
        {
            //进行INS转换
            var DrAccount = "";
            var DrMktSegment = "";
            var DrFolder = "";
            var DrBV = "";
            var results = from rs in dsMain.ITEMMapping
                          where rs.Item == "Installation"
                          select rs; ;
            foreach (var atemp in results)
            {
                //DrAccount = atemp.DrAccount;
                DrMktSegment = atemp.DrMktSegment;
                //DrFolder = atemp.DrFolder;
                DrBV = atemp.DrBV;
                break;
            }

            var CrAccount = "";
            var CrMktSegment = "";
            var CrFolder = "";
            var CrBV = "";
            var CrResults = from rs in dsMain.ITEMMapping
                            where rs.Item == "Installation"
                            select rs; ;
            foreach (var atemp in CrResults)
            {
                //CrAccount = atemp.CrAccount;
                CrMktSegment = atemp.CrMktSegment;
               //CrFolder = atemp.CrGETDFolder;
                CrBV = atemp.CrBV;
                break;
            }

            pbMain.Value = 0;
            pbMain.Maximum = uGridInput.Rows.Count;
            var iCount = 0;
            for (var i = 0; i < uGridInput.Rows.Count; i++)
            {
                //获取INS数据
                decimal dins;
                var cins = uGridInput.Rows[i].Cells["INS"].Value.ToString();
                var bins = decimal.TryParse(cins, out dins);
                if (!bins)
                    continue;
                if (dins == 0)
                    continue;
                //先写借
                WriteINSDr(i, dins, DrMktSegment, DrBV);
                WriteINSCr(i, dins, CrMktSegment, CrBV);
                pbMain.Value = i;
                iCount = iCount + 2;
            }
            return iCount;

        }
        /// <summary>
        /// INS输出写借方
        /// </summary>
        /// <param name="i"></param>
        private void WriteINSDr(int i, decimal iIns, string DrMktSegment, string DrBV)
        {
            var citem = "Installation";
            var LE = uGridInput.Rows[i].Cells["LE"].Value.ToString();
            var PL = uGridInput.Rows[i].Cells["PL"].Value.ToString();
            var Mod = uGridInput.Rows[i].Cells["MOD_Code"].Value.ToString();
            var DrAccount = "561000013";
            var DrCC = "";
            var DrFolder = "0000000";


            var results = from rs in dsMain.SDMatrix
                          where rs.Item == "Installation" && rs.PL == PL && rs.LE == LE && rs.ModCode == Mod
                          select rs; ;

            foreach (var atemp in results)
            {
                DrAccount = atemp.DrAccount;
                DrCC = atemp.DrCC;
                DrFolder = atemp.DrFolder;
                break;
            }

            var outputRow = dsMain.DataOutPut.NewDataOutPutRow();
            outputRow.ACCOUNTING_DATE = txtACCOUNTING_DATE.Text;
            //判断币种
            var cCurrency = "";

            if (LE.Equals("760110"))
            {
                outputRow.CURRENCY_CODE = "CNY";
            }
            else if (LE.Equals("700110"))
            {
                outputRow.CURRENCY_CODE = "USD";
            }

            outputRow.DATE_CREATED = "";
            outputRow.CURRENCY_CONV_DATE = "";
            outputRow.CURRENCY_CONV_TYPE = "";
            outputRow.CURRENCY_CONV_RATE = "";
            //Company
            outputRow.AFF_COMPANY = LE;
            outputRow.AFF_ACCOUNT = DrAccount;
            //成本中心
            outputRow.AFF_CENTER = DrCC;


            outputRow.AFF_BASE_VAR = DrBV;
            outputRow.AFF_MODALITY = Mod;
            outputRow.AFF_MKT_SEGMENT = DrMktSegment;
            //outputRow.AFF_FOLDER = DrFolder;

            //判断AFF_FOLDER
            outputRow.AFF_FOLDER = DrFolder;
            //判断Source
            if (LE.Length > 2)
            {
                if (LE.Substring(0, 2).Equals("76"))
                {
                    outputRow.AFF_SOURCE = "7601";
                }
                else if (LE.Substring(0, 2).Equals("70"))
                {
                    outputRow.AFF_SOURCE = "7001";
                }
            }
            //判断AFF_DESTINATION
            if (LE.Length > 2)
            {
                if (LE.Substring(0, 2).Equals("76"))
                {
                    outputRow.AFF_DESTINATION = "76";
                }
                else if (LE.Substring(0, 2).Equals("70"))
                {
                    outputRow.AFF_DESTINATION = "70";
                }
            }

            decimal dMoney = 0;
            if (LE.Equals("760110"))
            {
                dMoney = Math.Round(iIns * dMor, 2);
            }
            else if (LE.Equals("700110"))
            {
                dMoney = iIns;
            }

            outputRow.ENTERED_DR = dMoney.ToString();
            outputRow.ENTERED_CR = "";
            outputRow.ACCOUNTED_DR = "";
            outputRow.ACCOUNTED_CR = "";
            //判断SET_OF_BOOKS_ID
            if (LE.Length > 2)
            {
                if (LE.Substring(0, 2).Equals("76"))
                {
                    outputRow.SET_OF_BOOKS_ID = "050";
                }
                else if (LE.Substring(0, 2).Equals("70"))
                {
                    outputRow.SET_OF_BOOKS_ID = "051";
                }
            }
            outputRow.ACTUAL_FLAG = "A";
            outputRow.AFF_JOURNAL_CATEGORY = "Adjustment";
            outputRow.AFF_JOURNAL_SOURCE = "Spreadsheet";
            outputRow.PERIOD = "";
            outputRow.CODE_COMBINATION_ID = "";
            outputRow.COMPANY_CODE_MAP = "";
            outputRow.JOURNAL_SOURCE_MAP = "";
            outputRow.JOURNAL_CATEGORY_MAP = "";
            outputRow.JOURNAL_BATCH = txtJOURNAL_BATCH.Text;
            outputRow.BATCH_DESCRIPTION = txtBATCH_DESCRIPTION.Text;
            outputRow.JOURNAL_NAME = txtJOURNAL_BATCH.Text;
            outputRow.JOURNAL_DESCRIPTION = txtBATCH_DESCRIPTION.Text;
            outputRow.JOURNAL_REFERENCE = "";
            outputRow.JOURNALLINEDESC = txtDESC_INS.Text;
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
            outputRow.ORDER_NUM = uGridInput.Rows[i].Cells["OrderNumber"].Value.ToString();
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
        /// 输出写贷方
        /// </summary>
        /// <param name="i"></param>
        private void WriteINSCr(int i, decimal iIns, string CrMktSegment, string CrBV)
        {

            var citem = "Installation";
            var LE = uGridInput.Rows[i].Cells["LE"].Value.ToString();
            var PL = uGridInput.Rows[i].Cells["PL"].Value.ToString();
            var Mod = uGridInput.Rows[i].Cells["MOD_Code"].Value.ToString();
            var CrAccount = "561000013";
            var CrCC = "";
            var CrFolder = "0000000";
            var CrModCode = "";
            var CrPL = "";

            var results = from rs in dsMain.SDMatrix
                          where rs.Item == "Installation" && rs.PL == PL && rs.LE == LE && rs.ModCode == Mod
                          select rs;
            foreach (var atemp in results)
            {
                CrAccount = atemp.CrAccount;
                CrCC = atemp.CrCC;
                CrFolder = atemp.CrFolder;
                CrModCode = atemp.CrModCode;
                CrPL = atemp.CrPL;
                break;
            }


            var outputRow = dsMain.DataOutPut.NewDataOutPutRow();
            outputRow.ACCOUNTING_DATE = txtACCOUNTING_DATE.Text;
            //判断币种
            var cCurrency = "";

            if (LE.Equals("760110"))
            {
                outputRow.CURRENCY_CODE = "CNY";
            }
            else if (LE.Equals("700110"))
            {
                outputRow.CURRENCY_CODE = "USD";
            }

            outputRow.DATE_CREATED = "";
            outputRow.CURRENCY_CONV_DATE = "";
            outputRow.CURRENCY_CONV_TYPE = "";
            outputRow.CURRENCY_CONV_RATE = "";
            //Company
            outputRow.AFF_COMPANY = LE;
            outputRow.AFF_ACCOUNT = CrAccount;
            //成本中心
            outputRow.AFF_CENTER = CrCC;

            outputRow.AFF_BASE_VAR = CrBV;
            outputRow.AFF_MODALITY = CrModCode;
            outputRow.AFF_MKT_SEGMENT = CrMktSegment;
            outputRow.AFF_FOLDER = CrFolder;

            //判断AFF_FOLDER
            outputRow.AFF_FOLDER = CrFolder;
            //判断Source
            if (LE.Length > 2)
            {
                if (LE.Substring(0, 2).Equals("76"))
                {
                    outputRow.AFF_SOURCE = "7601";
                }
                else if (LE.Substring(0, 2).Equals("70"))
                {
                    outputRow.AFF_SOURCE = "7001";
                }
            }
            //判断AFF_DESTINATION
            if (LE.Length > 2)
            {
                if (LE.Substring(0, 2).Equals("76"))
                {
                    outputRow.AFF_DESTINATION = "76";
                }
                else if (LE.Substring(0, 2).Equals("70"))
                {
                    outputRow.AFF_DESTINATION = "70";
                }
            }

            decimal dMoney = 0;
            if (LE.Equals("760110"))
            {
                dMoney = Math.Round(iIns * dMor, 2);
            }
            else if (LE.Equals("700110"))
            {
                dMoney = iIns;
            }

            outputRow.ENTERED_DR = "";
            outputRow.ENTERED_CR = dMoney.ToString();
            outputRow.ACCOUNTED_DR = "";
            outputRow.ACCOUNTED_CR = "";
            //判断SET_OF_BOOKS_ID
            if (LE.Length > 2)
            {
                if (LE.Substring(0, 2).Equals("76"))
                {
                    outputRow.SET_OF_BOOKS_ID = "050";
                }
                else if (LE.Substring(0, 2).Equals("70"))
                {
                    outputRow.SET_OF_BOOKS_ID = "051";
                }
            }
            outputRow.ACTUAL_FLAG = "A";
            outputRow.AFF_JOURNAL_CATEGORY = "Adjustment";
            outputRow.AFF_JOURNAL_SOURCE = "Spreadsheet";
            outputRow.PERIOD = "";
            outputRow.CODE_COMBINATION_ID = "";
            outputRow.COMPANY_CODE_MAP = "";
            outputRow.JOURNAL_SOURCE_MAP = "";
            outputRow.JOURNAL_CATEGORY_MAP = "";
            outputRow.JOURNAL_BATCH = txtJOURNAL_BATCH.Text;
            outputRow.BATCH_DESCRIPTION = txtBATCH_DESCRIPTION.Text;
            outputRow.JOURNAL_NAME = txtJOURNAL_BATCH.Text;
            outputRow.JOURNAL_DESCRIPTION = txtBATCH_DESCRIPTION.Text;
            outputRow.JOURNAL_REFERENCE = "";
            outputRow.JOURNALLINEDESC = txtDESC_INS.Text;
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
            outputRow.ORDER_NUM = uGridInput.Rows[i].Cells["OrderNumber"].Value.ToString();
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
        /// 生成Warranty数据
        /// </summary>
        /// <returns></returns>
        private int GenerateWAR()
        {
            //进行INS转换
            var DrAccount = "";
            var DrMktSegment = "";
            var DrFolder = "";
            var DrBV = "";
            var results = from rs in dsMain.ITEMMapping
                          where rs.Item == "Warranty"
                          select rs; ;
            foreach (var atemp in results)
            {
                //DrAccount = atemp.DrAccount;
                DrMktSegment = atemp.DrMktSegment;
                //DrFolder = atemp.DrFolder;
                DrBV = atemp.DrBV;
                break;
            }

            var CrAccount = "";
            var CrMktSegment = "";
            var CrFolder = "";
            var CrBV = "";
            var CrResults = from rs in dsMain.ITEMMapping
                            where rs.Item == "Warranty"
                            select rs; ;
            foreach (var atemp in CrResults)
            {
                //CrAccount = atemp.CrAccount;
                CrMktSegment = atemp.CrMktSegment;
                //CrFolder = atemp.CrGETDFolder;
                CrBV = atemp.CrBV;
                break;
            }

            pbMain.Value = 0;
            pbMain.Maximum = uGridInput.Rows.Count;
            var iCount = 0;
            for (var i = 0; i < uGridInput.Rows.Count; i++)
            {
                //获取INS数据
                decimal dwar;
                var cwar = uGridInput.Rows[i].Cells["WAR"].Value.ToString();
                var bwar = decimal.TryParse(cwar, out dwar);
                if (!bwar)
                    continue;
                if (dwar == 0)
                    continue;
                //先写借
                WriteWARDr(i, dwar, DrMktSegment, DrBV);
                WriteWARCr(i, dwar, CrMktSegment, CrBV);
                pbMain.Value = i;
                iCount = iCount + 2;
            }
            return iCount;

        }
        /// <summary>
        /// WAR输出写借方
        /// </summary>
        /// <param name="i"></param>
        private void WriteWARDr(int i, decimal iIns, string DrMktSegment, string DrBV)
        {
            var citem = "Warranty";
            var LE = uGridInput.Rows[i].Cells["LE"].Value.ToString();
            var PL = uGridInput.Rows[i].Cells["PL"].Value.ToString();
            var Mod = uGridInput.Rows[i].Cells["MOD_Code"].Value.ToString();
            var DrAccount = "";
            var DrCC = "";
            var DrFolder = "";
            var results = from rs in dsMain.SDMatrix
                          where rs.Item == "Warranty" && rs.PL == PL && rs.LE == LE && rs.ModCode == Mod
                          select rs; ;

            foreach (var atemp in results)
            {
                DrAccount = atemp.DrAccount;
                DrCC = atemp.DrCC;
                DrFolder = atemp.DrFolder;
                break;
            }

            var outputRow = dsMain.DataOutPut.NewDataOutPutRow();
            outputRow.ACCOUNTING_DATE = txtACCOUNTING_DATE.Text;
            //判断币种
            var cCurrency = "";

            if (LE.Equals("760110"))
            {
                outputRow.CURRENCY_CODE = "CNY";
            }
            else if (LE.Equals("700110"))
            {
                outputRow.CURRENCY_CODE = "USD";
            }

            outputRow.DATE_CREATED = "";
            outputRow.CURRENCY_CONV_DATE = "";
            outputRow.CURRENCY_CONV_TYPE = "";
            outputRow.CURRENCY_CONV_RATE = "";
            //Company
            outputRow.AFF_COMPANY = LE;
            outputRow.AFF_ACCOUNT = DrAccount;
            //成本中心
            outputRow.AFF_CENTER = DrCC;


            outputRow.AFF_BASE_VAR = DrBV;
            outputRow.AFF_MODALITY = uGridInput.Rows[i].Cells["MOD_Code"].Value.ToString();
            outputRow.AFF_MKT_SEGMENT = DrMktSegment;
            outputRow.AFF_FOLDER = DrFolder;

            //判断AFF_FOLDER
            outputRow.AFF_FOLDER = DrFolder;
            //判断Source
            if (LE.Length > 2)
            {
                if (LE.Substring(0, 2).Equals("76"))
                {
                    outputRow.AFF_SOURCE = "7601";
                }
                else if (LE.Substring(0, 2).Equals("70"))
                {
                    outputRow.AFF_SOURCE = "7001";
                }
            }
            //判断AFF_DESTINATION
            if (LE.Length > 2)
            {
                if (LE.Substring(0, 2).Equals("76"))
                {
                    outputRow.AFF_DESTINATION = "76";
                }
                else if (LE.Substring(0, 2).Equals("70"))
                {
                    outputRow.AFF_DESTINATION = "70";
                }
            }
            decimal dMoney = 0;
            if (LE.Equals("760110"))
            {
                dMoney = Math.Round(iIns * dMor, 2);
            }
            else if (LE.Equals("700110"))
            {
                dMoney = iIns;
            }
            outputRow.ENTERED_DR = dMoney.ToString();
            outputRow.ENTERED_CR = "";
            outputRow.ACCOUNTED_DR = "";
            outputRow.ACCOUNTED_CR = "";
            //判断SET_OF_BOOKS_ID
            if (LE.Length > 2)
            {
                if (LE.Substring(0, 2).Equals("76"))
                {
                    outputRow.SET_OF_BOOKS_ID = "050";
                }
                else if (LE.Substring(0, 2).Equals("70"))
                {
                    outputRow.SET_OF_BOOKS_ID = "051";
                }
            }
            outputRow.ACTUAL_FLAG = "A";
            outputRow.AFF_JOURNAL_CATEGORY = "Adjustment";
            outputRow.AFF_JOURNAL_SOURCE = "Spreadsheet";
            outputRow.PERIOD = "";
            outputRow.CODE_COMBINATION_ID = "";
            outputRow.COMPANY_CODE_MAP = "";
            outputRow.JOURNAL_SOURCE_MAP = "";
            outputRow.JOURNAL_CATEGORY_MAP = "";
            outputRow.JOURNAL_BATCH = txtJOURNAL_BATCH.Text;
            outputRow.BATCH_DESCRIPTION = txtBATCH_DESCRIPTION.Text;
            outputRow.JOURNAL_NAME = txtJOURNAL_BATCH.Text;
            outputRow.JOURNAL_DESCRIPTION = txtBATCH_DESCRIPTION.Text;
            outputRow.JOURNAL_REFERENCE = "";
            outputRow.JOURNALLINEDESC = txtDESC_WAR.Text;
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
            outputRow.ORDER_NUM = uGridInput.Rows[i].Cells["OrderNumber"].Value.ToString();
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
        /// WAR输出写贷方
        /// </summary>
        /// <param name="i"></param>
        private void WriteWARCr(int i, decimal iIns, string CrMktSegment, string CrBV)
        {
            var citem = "Warranty";
            var LE = uGridInput.Rows[i].Cells["LE"].Value.ToString();
            var PL = uGridInput.Rows[i].Cells["PL"].Value.ToString();
            var Mod = uGridInput.Rows[i].Cells["MOD_Code"].Value.ToString();
            var CrAccount = "561000013";
            var CrCC = "";
            var CrFolder = "0000000";

            var CrModCode = "";
            var CrPL = "";
            var results = from rs in dsMain.SDMatrix
                          where rs.Item == "Warranty" && rs.PL == PL && rs.LE == LE && rs.ModCode == Mod
                          select rs;

            foreach (var atemp in results)
            {
                CrAccount = atemp.CrAccount;
                CrCC = atemp.CrCC;
                CrFolder = atemp.CrFolder;
                CrModCode = atemp.CrModCode;
                CrPL = atemp.CrPL;
                break;
            }
            var outputRow = dsMain.DataOutPut.NewDataOutPutRow();
            outputRow.ACCOUNTING_DATE = txtACCOUNTING_DATE.Text;
            //判断币种
            var cCurrency = "";

            if (LE.Equals("760110"))
            {
                outputRow.CURRENCY_CODE = "CNY";
            }
            else if (LE.Equals("700110"))
            {
                outputRow.CURRENCY_CODE = "USD";
            }

            outputRow.DATE_CREATED = "";
            outputRow.CURRENCY_CONV_DATE = "";
            outputRow.CURRENCY_CONV_TYPE = "";
            outputRow.CURRENCY_CONV_RATE = "";
            //Company
            outputRow.AFF_COMPANY = LE;
            outputRow.AFF_ACCOUNT = CrAccount;
            //成本中心
            outputRow.AFF_CENTER = CrCC;


            outputRow.AFF_BASE_VAR = CrBV;
            outputRow.AFF_MODALITY = CrModCode;
            outputRow.AFF_MKT_SEGMENT = CrMktSegment;
            outputRow.AFF_FOLDER = CrFolder;

            //判断AFF_FOLDER
            //if (LE.Length > 2)
            //{
            //    if (LE.Substring(0, 2).Equals("76"))
            //    {
            //        outputRow.AFF_FOLDER = "7081481";
            //    }
            //    else if (LE.Substring(0, 2).Equals("70"))
            //    {
            //        outputRow.AFF_FOLDER = "7081481";
            //    }
            //}
            //判断Source
            if (LE.Length > 2)
            {
                if (LE.Substring(0, 2).Equals("76"))
                {
                    outputRow.AFF_SOURCE = "7601";
                }
                else if (LE.Substring(0, 2).Equals("70"))
                {
                    outputRow.AFF_SOURCE = "7001";
                }
            }
            //判断AFF_DESTINATION
            if (LE.Length > 2)
            {
                if (LE.Substring(0, 2).Equals("76"))
                {
                    outputRow.AFF_DESTINATION = "76";
                }
                else if (LE.Substring(0, 2).Equals("70"))
                {
                    outputRow.AFF_DESTINATION = "70";
                }
            }


            decimal dMoney = 0;
            if (LE.Equals("760110"))
            {
                dMoney = Math.Round(iIns * dMor, 2);
            }
            else if (LE.Equals("700110"))
            {
                dMoney = iIns;
            }

            outputRow.ENTERED_DR = "";
            outputRow.ENTERED_CR = dMoney.ToString();
            outputRow.ACCOUNTED_DR = "";
            outputRow.ACCOUNTED_CR = "";
            //判断SET_OF_BOOKS_ID
            if (LE.Length > 2)
            {
                if (LE.Substring(0, 2).Equals("76"))
                {
                    outputRow.SET_OF_BOOKS_ID = "050";
                }
                else if (LE.Substring(0, 2).Equals("70"))
                {
                    outputRow.SET_OF_BOOKS_ID = "051";
                }
            }
            outputRow.ACTUAL_FLAG = "A";
            outputRow.AFF_JOURNAL_CATEGORY = "Adjustment";
            outputRow.AFF_JOURNAL_SOURCE = "Spreadsheet";
            outputRow.PERIOD = "";
            outputRow.CODE_COMBINATION_ID = "";
            outputRow.COMPANY_CODE_MAP = "";
            outputRow.JOURNAL_SOURCE_MAP = "";
            outputRow.JOURNAL_CATEGORY_MAP = "";
            outputRow.JOURNAL_BATCH = txtJOURNAL_BATCH.Text;
            outputRow.BATCH_DESCRIPTION = txtBATCH_DESCRIPTION.Text;
            outputRow.JOURNAL_NAME = txtJOURNAL_BATCH.Text;
            outputRow.JOURNAL_DESCRIPTION = txtBATCH_DESCRIPTION.Text;
            outputRow.JOURNAL_REFERENCE = "";
            outputRow.JOURNALLINEDESC = txtDESC_WAR.Text;
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
            outputRow.ORDER_NUM = uGridInput.Rows[i].Cells["OrderNumber"].Value.ToString();
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
        /// 生成APPI数据
        /// </summary>
        /// <returns></returns>
        private int GenerateAPPI()
        {
            //进行INS转换
            var DrAccount = "";
            var DrMktSegment = "";
            var DrFolder = "";
            var DrBV = "";
            var results = from rs in dsMain.ITEMMapping
                          where rs.Item == "Application"
                          select rs; ;
            foreach (var atemp in results)
            {
                //DrAccount = atemp.DrAccount;
                DrMktSegment = atemp.DrMktSegment;
                //DrFolder = atemp.DrFolder;
                DrBV = atemp.DrBV;
                break;
            }

            var CrAccount = "";
            var CrMktSegment = "";
            var CrFolder = "";
            var CrBV = "";
            var CrResults = from rs in dsMain.ITEMMapping
                            where rs.Item == "Application"
                            select rs; ;
            foreach (var atemp in CrResults)
            {
                //CrAccount = atemp.CrAccount;
                CrMktSegment = atemp.CrMktSegment;
                //CrFolder = atemp.CrGETDFolder;
                CrBV = atemp.CrBV;
                break;
            }

            pbMain.Value = 0;
            pbMain.Maximum = uGridInput.Rows.Count;
            var iCount = 0;
            for (var i = 0; i < uGridInput.Rows.Count; i++)
            {
                //获取INS数据
                decimal dwar;
                var cwar = uGridInput.Rows[i].Cells["APPI"].Value.ToString();
                var cModCode = uGridInput.Rows[i].Cells["MOD_Code"].Value.ToString();
                var bwar = decimal.TryParse(cwar, out dwar);
                if (!bwar)
                    continue;
                if (dwar == 0)
                    continue;
                //先写借
                if (!"751,731,781,721,907,841".Contains(cModCode))
                {
                    WriteAPPIDr(i, dwar, DrAccount, DrMktSegment, DrFolder, DrBV);
                    WriteAPPICr(i, dwar, CrAccount, CrMktSegment, CrFolder, CrBV);
                    iCount = iCount + 2;
                }

                pbMain.Value = i;

            }
            for (var i = 0; i < uGridInput.Rows.Count; i++)
            {
                //获取INS数据
                decimal dwar;
                var cwar = uGridInput.Rows[i].Cells["APPI"].Value.ToString();
                var cModCode = uGridInput.Rows[i].Cells["MOD_Code"].Value.ToString();
                var bwar = decimal.TryParse(cwar, out dwar);
                if (!bwar)
                    continue;
                if (dwar == 0)
                    continue;
                //先写借
                if ("751,731,781,721,907,841".Contains(cModCode))
                {
                    WriteAPPIDr(i, dwar, DrAccount, DrMktSegment, DrFolder, DrBV);
                    iCount = iCount + 1;
                }
                pbMain.Value = i;

            }
            //计算751的合计
            APPIGroupBy(CrAccount, CrMktSegment, CrFolder, CrBV);
            ////计算731的合计
            //iCount = iCount + APPIGroupBy(CrAccount, CrMktSegment, CrFolder, CrBV, "731");
            ////计算781的合计
            //iCount = iCount + APPIGroupBy(CrAccount, CrMktSegment, CrFolder, CrBV, "781");
            ////计算721的合计
            //iCount = iCount + APPIGroupBy(CrAccount, CrMktSegment, CrFolder, CrBV, "721");
            ////计算907的合计
            //iCount = iCount + APPIGroupBy(CrAccount, CrMktSegment, CrFolder, CrBV, "907");
            ////计算841的合计
            //iCount = iCount + APPIGroupBy(CrAccount, CrMktSegment, CrFolder, CrBV, "841");



            return iCount;

        }
        /// <summary>
        /// 计算APPI的Groupby 结果
        /// </summary>
        /// <param name="CrAccount"></param>
        /// <param name="CrMktSegment"></param>
        /// <param name="CrFolder"></param>
        /// <param name="CrBV"></param>
        private int APPIGroupBy(string CrAccount, string CrMktSegment, string CrFolder, string CrBV)
        {
            string[] strModCode = new string[6] { "751", "731", "781", "721", "907", "841" };
            var resultAPPI = from u in dsMain.DataInput
                             where strModCode.Contains(u.MOD_Code)
                             group u by new { LE = u.LE} into g
                             select new
                             {
                                 g.Key.LE,
                                 APPI = g.Sum(c => c.APPI)
                             };
            var iCount = 0;
            foreach (var itemAPPI in resultAPPI)
            {
                if (itemAPPI.APPI != 0)
                {
                    WriteAPPICrGoupBy(itemAPPI.LE,"990", itemAPPI.APPI, CrAccount, CrMktSegment, CrFolder, CrBV);
                    iCount = iCount + 1;
                }
            }
            return iCount;
        }
        /// <summary>
        /// APPI输出写借方
        /// </summary>
        /// <param name="i"></param>
        private void WriteAPPIDr(int i, decimal iIns, string DrAccount, string DrMktSegment, string DrFolder, string DrBV)
        {

            var citem = "Application";
            var LE = uGridInput.Rows[i].Cells["LE"].Value.ToString();
            var PL = uGridInput.Rows[i].Cells["PL"].Value.ToString();
            var Mod = uGridInput.Rows[i].Cells["MOD_Code"].Value.ToString();

            var DrCC = "";
            var results = from rs in dsMain.SDMatrix
                          where rs.Item == "Application" && rs.PL == PL && rs.LE == LE && rs.ModCode == Mod
                          select rs; ;

            foreach (var atemp in results)
            {
                DrAccount = atemp.DrAccount;
                DrCC = atemp.DrCC;
                DrFolder = atemp.DrFolder;
                break;
            }

            var outputRow = dsMain.DataOutPut.NewDataOutPutRow();
            outputRow.ACCOUNTING_DATE = txtACCOUNTING_DATE.Text;
            //判断币种
            var cCurrency = "";

            if (LE.Equals("760110"))
            {
                outputRow.CURRENCY_CODE = "CNY";
            }
            else if (LE.Equals("700110"))
            {
                outputRow.CURRENCY_CODE = "USD";
            }

            outputRow.DATE_CREATED = "";
            outputRow.CURRENCY_CONV_DATE = "";
            outputRow.CURRENCY_CONV_TYPE = "";
            outputRow.CURRENCY_CONV_RATE = "";
            //Company
            outputRow.AFF_COMPANY = LE;
            outputRow.AFF_ACCOUNT = DrAccount;
            //成本中心
            outputRow.AFF_CENTER = DrCC;

            outputRow.AFF_BASE_VAR = DrBV;
            outputRow.AFF_MODALITY = uGridInput.Rows[i].Cells["MOD_Code"].Value.ToString();
            outputRow.AFF_MKT_SEGMENT = DrMktSegment;
            outputRow.AFF_FOLDER = DrFolder;

            //判断AFF_FOLDER
            //if (LE.Length > 2)
            //{
            //    if (LE.Substring(0, 2).Equals("76"))
            //    {
            //        outputRow.AFF_FOLDER = "7081481";
            //    }
            //    else if (LE.Substring(0, 2).Equals("70"))
            //    {
            //        outputRow.AFF_FOLDER = "7081481";
            //    }
            //}
            //判断Source
            if (LE.Length > 2)
            {
                if (LE.Substring(0, 2).Equals("76"))
                {
                    outputRow.AFF_SOURCE = "7601";
                }
                else if (LE.Substring(0, 2).Equals("70"))
                {
                    outputRow.AFF_SOURCE = "7001";
                }
            }
            //判断AFF_DESTINATION
            if (LE.Length > 2)
            {
                if (LE.Substring(0, 2).Equals("76"))
                {
                    outputRow.AFF_DESTINATION = "76";
                }
                else if (LE.Substring(0, 2).Equals("70"))
                {
                    outputRow.AFF_DESTINATION = "70";
                }
            }
            decimal dMoney = 0;
            if (LE.Equals("760110"))
            {
                dMoney = Math.Round(iIns * dMor, 2);
            }
            else if (LE.Equals("700110"))
            {
                dMoney = iIns;
            }

            outputRow.ENTERED_DR = dMoney.ToString();
            outputRow.ENTERED_CR = "";
            outputRow.ACCOUNTED_DR = "";
            outputRow.ACCOUNTED_CR = "";
            //判断SET_OF_BOOKS_ID
            if (LE.Length > 2)
            {
                if (LE.Substring(0, 2).Equals("76"))
                {
                    outputRow.SET_OF_BOOKS_ID = "050";
                }
                else if (LE.Substring(0, 2).Equals("70"))
                {
                    outputRow.SET_OF_BOOKS_ID = "051";
                }
            }
            outputRow.ACTUAL_FLAG = "A";
            outputRow.AFF_JOURNAL_CATEGORY = "Adjustment";
            outputRow.AFF_JOURNAL_SOURCE = "Spreadsheet";
            outputRow.PERIOD = "";
            outputRow.CODE_COMBINATION_ID = "";
            outputRow.COMPANY_CODE_MAP = "";
            outputRow.JOURNAL_SOURCE_MAP = "";
            outputRow.JOURNAL_CATEGORY_MAP = "";
            outputRow.JOURNAL_BATCH = txtJOURNAL_BATCH.Text;
            outputRow.BATCH_DESCRIPTION = txtBATCH_DESCRIPTION.Text;
            outputRow.JOURNAL_NAME = txtJOURNAL_BATCH.Text;
            outputRow.JOURNAL_DESCRIPTION = txtBATCH_DESCRIPTION.Text;
            outputRow.JOURNAL_REFERENCE = "";
            outputRow.JOURNALLINEDESC = txtDESC_APPI.Text;
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
            outputRow.ORDER_NUM = uGridInput.Rows[i].Cells["OrderNumber"].Value.ToString();
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
        /// APPI输出写贷方
        /// </summary>
        /// <param name="i"></param>
        private void WriteAPPICr(int i, decimal iIns, string CrAccount, string CrMktSegment, string CrFolder, string CrBV)
        {

            var citem = "Application";
            var LE = uGridInput.Rows[i].Cells["LE"].Value.ToString();
            var PL = uGridInput.Rows[i].Cells["PL"].Value.ToString();
            var Mod = uGridInput.Rows[i].Cells["MOD_Code"].Value.ToString();
            CrAccount = "";
            var CrCC = "";
            CrFolder = "";
            var CrModCode = "";
            var results = from rs in dsMain.SDMatrix
                          where rs.Item == "Application" && rs.PL == PL && rs.LE == LE && rs.ModCode == Mod
                          select rs; ;

            foreach (var atemp in results)
            {
                CrAccount = atemp.CrAccount;
                CrCC = atemp.CrCC;
                CrFolder = atemp.CrFolder;
                CrModCode = atemp.CrModCode;
                break;
            }
            var outputRow = dsMain.DataOutPut.NewDataOutPutRow();
            outputRow.ACCOUNTING_DATE = txtACCOUNTING_DATE.Text;
            //判断币种
            var cCurrency = "";

            if (LE.Equals("760110"))
            {
                outputRow.CURRENCY_CODE = "CNY";
            }
            else if (LE.Equals("700110"))
            {
                outputRow.CURRENCY_CODE = "USD";
            }

            outputRow.DATE_CREATED = "";
            outputRow.CURRENCY_CONV_DATE = "";
            outputRow.CURRENCY_CONV_TYPE = "";
            outputRow.CURRENCY_CONV_RATE = "";
            //Company
            outputRow.AFF_COMPANY = LE;
            outputRow.AFF_ACCOUNT = CrAccount;
            //成本中心
            outputRow.AFF_CENTER = CrCC;


            outputRow.AFF_BASE_VAR = CrBV;
            outputRow.AFF_MODALITY = CrModCode;
            outputRow.AFF_MKT_SEGMENT = CrMktSegment;
            outputRow.AFF_FOLDER = CrFolder;

            //判断AFF_FOLDER
            //if (LE.Length > 2)
            //{
            //    if (LE.Substring(0, 2).Equals("76"))
            //    {
            //        outputRow.AFF_FOLDER = "7081481";
            //    }
            //    else if (LE.Substring(0, 2).Equals("70"))
            //    {
            //        outputRow.AFF_FOLDER = "7081481";
            //    }
            //}
            //判断Source
            if (LE.Length > 2)
            {
                if (LE.Substring(0, 2).Equals("76"))
                {
                    outputRow.AFF_SOURCE = "7601";
                }
                else if (LE.Substring(0, 2).Equals("70"))
                {
                    outputRow.AFF_SOURCE = "7001";
                }
            }
            //判断AFF_DESTINATION
            if (LE.Length > 2)
            {
                if (LE.Substring(0, 2).Equals("76"))
                {
                    outputRow.AFF_DESTINATION = "76";
                }
                else if (LE.Substring(0, 2).Equals("70"))
                {
                    outputRow.AFF_DESTINATION = "70";
                }
            }

            decimal dMoney = 0;
            if (LE.Equals("760110"))
            {
                dMoney = Math.Round(iIns * dMor, 2);
            }
            else if (LE.Equals("700110"))
            {
                dMoney = iIns;
            }

            outputRow.ENTERED_DR = "";
            outputRow.ENTERED_CR = dMoney.ToString();
            outputRow.ACCOUNTED_DR = "";
            outputRow.ACCOUNTED_CR = "";
            //判断SET_OF_BOOKS_ID
            if (LE.Length > 2)
            {
                if (LE.Substring(0, 2).Equals("76"))
                {
                    outputRow.SET_OF_BOOKS_ID = "050";
                }
                else if (LE.Substring(0, 2).Equals("70"))
                {
                    outputRow.SET_OF_BOOKS_ID = "051";
                }
            }
            outputRow.ACTUAL_FLAG = "A";
            outputRow.AFF_JOURNAL_CATEGORY = "Adjustment";
            outputRow.AFF_JOURNAL_SOURCE = "Spreadsheet";
            outputRow.PERIOD = "";
            outputRow.CODE_COMBINATION_ID = "";
            outputRow.COMPANY_CODE_MAP = "";
            outputRow.JOURNAL_SOURCE_MAP = "";
            outputRow.JOURNAL_CATEGORY_MAP = "";
            outputRow.JOURNAL_BATCH = txtJOURNAL_BATCH.Text;
            outputRow.BATCH_DESCRIPTION = txtBATCH_DESCRIPTION.Text;
            outputRow.JOURNAL_NAME = txtJOURNAL_BATCH.Text;
            outputRow.JOURNAL_DESCRIPTION = txtBATCH_DESCRIPTION.Text;
            outputRow.JOURNAL_REFERENCE = "";
            outputRow.JOURNALLINEDESC = txtDESC_APPI.Text;
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
            outputRow.ORDER_NUM = uGridInput.Rows[i].Cells["OrderNumber"].Value.ToString();
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
        /// 执行APPI的Groupy结果
        /// </summary>
        /// <param name="LE"></param>
        /// <param name="ModCode"></param>
        /// <param name="iIns"></param>
        /// <param name="CrAccount"></param>
        /// <param name="CrMktSegment"></param>
        /// <param name="CrFolder"></param>
        /// <param name="CrBV"></param>
        private void WriteAPPICrGoupBy(string LE, string ModCode, decimal iIns, string CrAccount, string CrMktSegment, string CrFolder, string CrBV)
        {

            var citem = "Application";
            var PL = "L23";

            var CrCC = "";

            var CrModCode = "";
            var results = from rs in dsMain.SDMatrix
                          where rs.Item == "Application" && rs.PL == PL && rs.LE == LE && rs.ModCode == ModCode
                          select rs;

            foreach (var atemp in results)
            {
                CrAccount = atemp.CrAccount;
                CrCC = atemp.CrCC;
                CrFolder = atemp.CrFolder;
                CrModCode = atemp.CrModCode;

                break;
            }
            var outputRow = dsMain.DataOutPut.NewDataOutPutRow();
            outputRow.ACCOUNTING_DATE = txtACCOUNTING_DATE.Text;
            //判断币种
            var cCurrency = "";
            if (LE.Equals("760110"))
            {
                outputRow.CURRENCY_CODE = "CNY";
            }
            else if (LE.Equals("700110"))
            {
                outputRow.CURRENCY_CODE = "USD";
            }

            outputRow.DATE_CREATED = "";
            outputRow.CURRENCY_CONV_DATE = "";
            outputRow.CURRENCY_CONV_TYPE = "";
            outputRow.CURRENCY_CONV_RATE = "";
            //Company
            outputRow.AFF_COMPANY = LE;
            outputRow.AFF_ACCOUNT = CrAccount;
            //成本中心
            outputRow.AFF_CENTER = CrCC;

            if (CrAccount.Equals("560000080"))
            {
                outputRow.AFF_BASE_VAR = "1";
            }
            else
            {
                outputRow.AFF_BASE_VAR = CrBV;
            }
            


            outputRow.AFF_MODALITY = CrModCode;

            if (CrAccount.Equals("560000080"))
            {

                outputRow.AFF_MKT_SEGMENT = "90";
            }
            else
            {
                outputRow.AFF_MKT_SEGMENT = CrMktSegment;
            }


            outputRow.AFF_FOLDER = CrFolder;

            //判断AFF_FOLDER
            //if (LE.Length > 2)
            //{
            //    if (LE.Substring(0, 2).Equals("76"))
            //    {
            //        outputRow.AFF_FOLDER = "7081481";
            //    }
            //    else if (LE.Substring(0, 2).Equals("70"))
            //    {
            //        outputRow.AFF_FOLDER = "7081481";
            //    }
            //}
            //判断Source
            if (LE.Length > 2)
            {
                if (LE.Substring(0, 2).Equals("76"))
                {
                    outputRow.AFF_SOURCE = "7601";
                }
                else if (LE.Substring(0, 2).Equals("70"))
                {
                    outputRow.AFF_SOURCE = "7001";
                }
            }
            //判断AFF_DESTINATION
            if (LE.Length > 2)
            {
                if (LE.Substring(0, 2).Equals("76"))
                {
                    outputRow.AFF_DESTINATION = "76";
                }
                else if (LE.Substring(0, 2).Equals("70"))
                {
                    outputRow.AFF_DESTINATION = "70";
                }
            }

            decimal dMoney = 0;
            if (LE.Equals("760110"))
            {
                dMoney = Math.Round(iIns * dMor, 2);
            }
            else if (LE.Equals("700110"))
            {
                dMoney = iIns;
            }

            outputRow.ENTERED_DR = "";
            outputRow.ENTERED_CR = dMoney.ToString();
            outputRow.ACCOUNTED_DR = "";
            outputRow.ACCOUNTED_CR = "";
            //判断SET_OF_BOOKS_ID
            if (LE.Length > 2)
            {
                if (LE.Substring(0, 2).Equals("76"))
                {
                    outputRow.SET_OF_BOOKS_ID = "050";
                }
                else if (LE.Substring(0, 2).Equals("70"))
                {
                    outputRow.SET_OF_BOOKS_ID = "051";
                }
            }
            outputRow.ACTUAL_FLAG = "A";
            outputRow.AFF_JOURNAL_CATEGORY = "Adjustment";
            outputRow.AFF_JOURNAL_SOURCE = "Spreadsheet";
            outputRow.PERIOD = "";
            outputRow.CODE_COMBINATION_ID = "";
            outputRow.COMPANY_CODE_MAP = "";
            outputRow.JOURNAL_SOURCE_MAP = "";
            outputRow.JOURNAL_CATEGORY_MAP = "";
            outputRow.JOURNAL_BATCH = txtJOURNAL_BATCH.Text;
            outputRow.BATCH_DESCRIPTION = txtBATCH_DESCRIPTION.Text;
            outputRow.JOURNAL_NAME = txtJOURNAL_BATCH.Text;
            outputRow.JOURNAL_DESCRIPTION = txtBATCH_DESCRIPTION.Text;
            outputRow.JOURNAL_REFERENCE = "";
            outputRow.JOURNALLINEDESC = txtDESC_APPI.Text;
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
        /// 生成APPII数据
        /// </summary>
        /// <returns></returns>
        private int APPIIGenerateAPPII()
        {
            //进行INS转换
            var DrAccount = "";
            var DrMktSegment = "";
            var DrFolder = "";
            var DrBV = "1";
            var results = from rs in dsMain.ITEMMapping
                          where rs.Item == "Adv App"
                          select rs; ;
            foreach (var atemp in results)
            {
                //DrAccount = atemp.DrAccount;
                DrMktSegment = atemp.DrMktSegment;
                //DrFolder = atemp.DrFolder;
                DrBV = atemp.DrBV;
                break;
            }

            var CrAccount = "";
            var CrMktSegment = "";
            var CrFolder = "";
            var CrBV = "";
            var CrResults = from rs in dsMain.ITEMMapping
                            where rs.Item == "Adv App"
                            select rs; ;
            foreach (var atemp in CrResults)
            {
                //CrAccount = atemp.CrAccount;
                CrMktSegment = atemp.CrMktSegment;
                //CrFolder = atemp.CrGETDFolder;
                CrBV = atemp.CrBV;
                break;
            }

            pbMain.Value = 0;
            pbMain.Maximum = uGridInput.Rows.Count;
            var iCount = 0;
            for (var i = 0; i < uGridInput.Rows.Count; i++)
            {
                //获取INS数据
                decimal dwar;
                var cwar = uGridInput.Rows[i].Cells["APPII"].Value.ToString();
                var cModCode = uGridInput.Rows[i].Cells["MOD_Code"].Value.ToString();
                var bwar = decimal.TryParse(cwar, out dwar);
                if (!bwar)
                    continue;
                if (dwar == 0)
                    continue;
                //先写借
                if (!"751,731,781,721,907,841".Contains(cModCode))
                {
                    WriteAPPIIDrAPPII(i, dwar, DrAccount, DrMktSegment, DrFolder, DrBV);
                    WriteAPPIICrAPPII(i, dwar, CrAccount, CrMktSegment, CrFolder, CrBV);
                    iCount = iCount + 2;
                }

                pbMain.Value = i;

            }
            for (var i = 0; i < uGridInput.Rows.Count; i++)
            {
                //获取INS数据
                decimal dwar;
                var cwar = uGridInput.Rows[i].Cells["APPII"].Value.ToString();
                var cModCode = uGridInput.Rows[i].Cells["MOD_Code"].Value.ToString();
                var bwar = decimal.TryParse(cwar, out dwar);
                if (!bwar)
                    continue;
                if (dwar == 0)
                    continue;
                //先写借
                if ("751,731,781,721,907,841".Contains(cModCode))
                {
                    WriteAPPIIDrAPPII(i, dwar, DrAccount, DrMktSegment, DrFolder, DrBV);
                    iCount = iCount + 1;
                }
                pbMain.Value = i;

            }
            //计算751的合计
            iCount = iCount + APPIIGroupByAPPII(CrAccount, CrMktSegment, CrFolder, CrBV, "751");
            //计算731的合计
            iCount = iCount + APPIIGroupByAPPII(CrAccount, CrMktSegment, CrFolder, CrBV, "731");
            //计算781的合计
            iCount = iCount + APPIIGroupByAPPII(CrAccount, CrMktSegment, CrFolder, CrBV, "781");
            //计算721的合计
            iCount = iCount + APPIIGroupByAPPII(CrAccount, CrMktSegment, CrFolder, CrBV, "721");
            //计算907的合计
            iCount = iCount + APPIIGroupByAPPII(CrAccount, CrMktSegment, CrFolder, CrBV, "907");
            //计算841的合计
            iCount = iCount + APPIIGroupByAPPII(CrAccount, CrMktSegment, CrFolder, CrBV, "841");


            return iCount;

        }
        /// <summary>
        /// 计算APPII的Groupby 结果
        /// </summary>
        /// <param name="CrAccount"></param>
        /// <param name="CrMktSegment"></param>
        /// <param name="CrFolder"></param>
        /// <param name="CrBV"></param>
        private int APPIIGroupByAPPII(string CrAccount, string CrMktSegment, string CrFolder, string CrBV, string ModCode)
        {
            var resultAPPI = from u in dsMain.DataInput
                             where u.MOD_Code == ModCode
                             group u by new { LE = u.LE, MOD_Code = u.MOD_Code } into g
                             select new
                             {
                                 g.Key.LE,
                                 g.Key.MOD_Code,
                                 APPI = g.Sum(c => c.APPII)
                             };
            var iCount = 0;
            foreach (var itemAPPI in resultAPPI)
            {
                if (itemAPPI.APPI != 0)
                {
                    WriteAPPIICrGoupByAPPII(itemAPPI.LE, itemAPPI.MOD_Code, itemAPPI.APPI, CrAccount, CrMktSegment, CrFolder, CrBV);
                    iCount = iCount + 1;
                }
            }
            return iCount;
        }
        /// <summary>
        /// APPII输出写借方
        /// </summary>
        /// <param name="i"></param>
        private void WriteAPPIIDrAPPII(int i, decimal iIns, string DrAccount, string DrMktSegment, string DrFolder, string DrBV)
        {

            var citem = "Adv App";
            var LE = uGridInput.Rows[i].Cells["LE"].Value.ToString();
            var PL = uGridInput.Rows[i].Cells["PL"].Value.ToString();
            var Mod = uGridInput.Rows[i].Cells["MOD_Code"].Value.ToString();

            var DrCC = "";
            var results = from rs in dsMain.SDMatrix
                          where rs.Item == "Adv App" && rs.PL == PL && rs.LE == LE && rs.ModCode == Mod
                          select rs; ;

            foreach (var atemp in results)
            {
                DrAccount = atemp.DrAccount;
                DrCC = atemp.DrCC;
                DrFolder = atemp.DrFolder;
                break;
            }

            var outputRow = dsMain.DataOutPut.NewDataOutPutRow();
            outputRow.ACCOUNTING_DATE = txtACCOUNTING_DATE.Text;
            //判断币种
            var cCurrency = "";

            if (LE.Equals("760110"))
            {
                outputRow.CURRENCY_CODE = "CNY";
            }
            else if (LE.Equals("700110"))
            {
                outputRow.CURRENCY_CODE = "USD";
            }

            outputRow.DATE_CREATED = "";
            outputRow.CURRENCY_CONV_DATE = "";
            outputRow.CURRENCY_CONV_TYPE = "";
            outputRow.CURRENCY_CONV_RATE = "";
            //Company
            outputRow.AFF_COMPANY = LE;
            outputRow.AFF_ACCOUNT = DrAccount;
            //成本中心

            outputRow.AFF_CENTER = DrCC;


            outputRow.AFF_BASE_VAR = DrBV;
            outputRow.AFF_MODALITY = uGridInput.Rows[i].Cells["MOD_Code"].Value.ToString();
            outputRow.AFF_MKT_SEGMENT = DrMktSegment;
            outputRow.AFF_FOLDER = DrFolder;

            //判断AFF_FOLDER
            //if (LE.Length > 2)
            //{
            //    if (LE.Substring(0, 2).Equals("76"))
            //    {
            //        outputRow.AFF_FOLDER = "7081481";
            //    }
            //    else if (LE.Substring(0, 2).Equals("70"))
            //    {
            //        outputRow.AFF_FOLDER = "7081481";
            //    }
            //}
            //判断Source
            if (LE.Length > 2)
            {
                if (LE.Substring(0, 2).Equals("76"))
                {
                    outputRow.AFF_SOURCE = "7601";
                }
                else if (LE.Substring(0, 2).Equals("70"))
                {
                    outputRow.AFF_SOURCE = "7001";
                }
            }
            //判断AFF_DESTINATION
            if (LE.Length > 2)
            {
                if (LE.Substring(0, 2).Equals("76"))
                {
                    outputRow.AFF_DESTINATION = "76";
                }
                else if (LE.Substring(0, 2).Equals("70"))
                {
                    outputRow.AFF_DESTINATION = "70";
                }
            }
            decimal dMoney = 0;
            if (LE.Equals("760110"))
            {
                dMoney = Math.Round(iIns * dMor, 2);
            }
            else if (LE.Equals("700110"))
            {
                dMoney = iIns;
            }
            outputRow.ENTERED_DR = dMoney.ToString();
            outputRow.ENTERED_CR = "";
            outputRow.ACCOUNTED_DR = "";
            outputRow.ACCOUNTED_CR = "";
            //判断SET_OF_BOOKS_ID
            if (LE.Length > 2)
            {
                if (LE.Substring(0, 2).Equals("76"))
                {
                    outputRow.SET_OF_BOOKS_ID = "050";
                }
                else if (LE.Substring(0, 2).Equals("70"))
                {
                    outputRow.SET_OF_BOOKS_ID = "051";
                }
            }
            outputRow.ACTUAL_FLAG = "A";
            outputRow.AFF_JOURNAL_CATEGORY = "Adjustment";
            outputRow.AFF_JOURNAL_SOURCE = "Spreadsheet";
            outputRow.PERIOD = "";
            outputRow.CODE_COMBINATION_ID = "";
            outputRow.COMPANY_CODE_MAP = "";
            outputRow.JOURNAL_SOURCE_MAP = "";
            outputRow.JOURNAL_CATEGORY_MAP = "";
            outputRow.JOURNAL_BATCH = txtJOURNAL_BATCH.Text;
            outputRow.BATCH_DESCRIPTION = txtBATCH_DESCRIPTION.Text;
            outputRow.JOURNAL_NAME = txtJOURNAL_BATCH.Text;
            outputRow.JOURNAL_DESCRIPTION = txtBATCH_DESCRIPTION.Text;
            outputRow.JOURNAL_REFERENCE = "";
            outputRow.JOURNALLINEDESC = txtDESC_APPII.Text;
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
            outputRow.ORDER_NUM = uGridInput.Rows[i].Cells["OrderNumber"].Value.ToString();
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
        /// APPII输出写贷方
        /// </summary>
        /// <param name="i"></param>
        private void WriteAPPIICrAPPII(int i, decimal iIns, string CrAccount, string CrMktSegment, string CrFolder, string CrBV)
        {
            var citem = "Adv App";
            var LE = uGridInput.Rows[i].Cells["LE"].Value.ToString();
            var PL = uGridInput.Rows[i].Cells["PL"].Value.ToString();
            var Mod = uGridInput.Rows[i].Cells["MOD_Code"].Value.ToString();
            CrAccount = "561000013";
            var CrCC = "";
            CrFolder = "0000000";

            var CrModCode = "";
            var results = from rs in dsMain.SDMatrix
                          where rs.Item == "Adv App" && rs.PL == PL && rs.LE == LE && rs.ModCode == Mod
                          select rs; ;

            foreach (var atemp in results)
            {
                CrAccount = atemp.CrAccount;
                CrCC = atemp.CrCC;
                CrFolder = atemp.CrFolder;
                CrModCode = atemp.CrModCode;
                break;
            }

            var outputRow = dsMain.DataOutPut.NewDataOutPutRow();
            outputRow.ACCOUNTING_DATE = txtACCOUNTING_DATE.Text;
            //判断币种
            var cCurrency = "";
            if (LE.Equals("760110"))
            {
                outputRow.CURRENCY_CODE = "CNY";
            }
            else if (LE.Equals("700110"))
            {
                outputRow.CURRENCY_CODE = "USD";
            }

            outputRow.DATE_CREATED = "";
            outputRow.CURRENCY_CONV_DATE = "";
            outputRow.CURRENCY_CONV_TYPE = "";
            outputRow.CURRENCY_CONV_RATE = "";
            //Company
            outputRow.AFF_COMPANY = LE;
            outputRow.AFF_ACCOUNT = CrAccount;
            //成本中心
            outputRow.AFF_CENTER = CrCC;
            outputRow.AFF_BASE_VAR = CrBV;
            outputRow.AFF_MODALITY = CrModCode;
            outputRow.AFF_MKT_SEGMENT = CrMktSegment;
            outputRow.AFF_FOLDER = CrFolder;

            //判断AFF_FOLDER
            //if (LE.Length > 2)
            //{
            //    if (LE.Substring(0, 2).Equals("76"))
            //    {
            //        outputRow.AFF_FOLDER = "7081481";
            //    }
            //    else if (LE.Substring(0, 2).Equals("70"))
            //    {
            //        outputRow.AFF_FOLDER = "7081481";
            //    }
            //}
            //判断Source
            if (LE.Length > 2)
            {
                if (LE.Substring(0, 2).Equals("76"))
                {
                    outputRow.AFF_SOURCE = "7601";
                }
                else if (LE.Substring(0, 2).Equals("70"))
                {
                    outputRow.AFF_SOURCE = "7001";
                }
            }
            //判断AFF_DESTINATION
            if (LE.Length > 2)
            {
                if (LE.Substring(0, 2).Equals("76"))
                {
                    outputRow.AFF_DESTINATION = "76";
                }
                else if (LE.Substring(0, 2).Equals("70"))
                {
                    outputRow.AFF_DESTINATION = "70";
                }
            }
            decimal dMoney = 0;
            if (LE.Equals("760110"))
            {
                dMoney = Math.Round(iIns * dMor, 2);
            }
            else if (LE.Equals("700110"))
            {
                dMoney = iIns;
            }

            outputRow.ENTERED_DR = "";
            outputRow.ENTERED_CR = dMoney.ToString();
            outputRow.ACCOUNTED_DR = "";
            outputRow.ACCOUNTED_CR = "";
            //判断SET_OF_BOOKS_ID
            if (LE.Length > 2)
            {
                if (LE.Substring(0, 2).Equals("76"))
                {
                    outputRow.SET_OF_BOOKS_ID = "050";
                }
                else if (LE.Substring(0, 2).Equals("70"))
                {
                    outputRow.SET_OF_BOOKS_ID = "051";
                }
            }
            outputRow.ACTUAL_FLAG = "A";
            outputRow.AFF_JOURNAL_CATEGORY = "Adjustment";
            outputRow.AFF_JOURNAL_SOURCE = "Spreadsheet";
            outputRow.PERIOD = "";
            outputRow.CODE_COMBINATION_ID = "";
            outputRow.COMPANY_CODE_MAP = "";
            outputRow.JOURNAL_SOURCE_MAP = "";
            outputRow.JOURNAL_CATEGORY_MAP = "";
            outputRow.JOURNAL_BATCH = txtJOURNAL_BATCH.Text;
            outputRow.BATCH_DESCRIPTION = txtBATCH_DESCRIPTION.Text;
            outputRow.JOURNAL_NAME = txtJOURNAL_BATCH.Text;
            outputRow.JOURNAL_DESCRIPTION = txtBATCH_DESCRIPTION.Text;
            outputRow.JOURNAL_REFERENCE = "";
            outputRow.JOURNALLINEDESC = txtDESC_APPII.Text;
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
            outputRow.ORDER_NUM = uGridInput.Rows[i].Cells["OrderNumber"].Value.ToString();
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
        /// 写APPII的Groupy后的贷方数据
        /// </summary>
        /// <param name="LE"></param>
        /// <param name="ModCode"></param>
        /// <param name="iIns"></param>
        /// <param name="CrAccount"></param>
        /// <param name="CrMktSegment"></param>
        /// <param name="CrFolder"></param>
        /// <param name="CrBV"></param>
        private void WriteAPPIICrGoupByAPPII(string LE, string ModCode, decimal iIns, string CrAccount, string CrMktSegment, string CrFolder, string CrBV)
        {
            var citem = "Adv App";

            var PL = "L23";
            var CrCC = "";
            var CrModCode = "";
            var results = from rs in dsMain.SDMatrix
                          where rs.Item == "Adv App" && rs.PL == PL && rs.LE == LE && rs.ModCode == ModCode
                          select rs; ;

            foreach (var atemp in results)
            {
                CrAccount = atemp.CrAccount;
                CrCC = atemp.CrCC;
                CrFolder = atemp.CrFolder;
                CrModCode = atemp.CrModCode;
                break;
            }
            var outputRow = dsMain.DataOutPut.NewDataOutPutRow();
            outputRow.ACCOUNTING_DATE = txtACCOUNTING_DATE.Text;
            //判断币种
            var cCurrency = "";
            if (LE.Equals("760110"))
            {
                outputRow.CURRENCY_CODE = "CNY";
            }
            else if (LE.Equals("700110"))
            {
                outputRow.CURRENCY_CODE = "USD";
            }

            outputRow.DATE_CREATED = "";
            outputRow.CURRENCY_CONV_DATE = "";
            outputRow.CURRENCY_CONV_TYPE = "";
            outputRow.CURRENCY_CONV_RATE = "";
            //Company
            outputRow.AFF_COMPANY = LE;
            outputRow.AFF_ACCOUNT = CrAccount;
            //成本中心
            outputRow.AFF_CENTER = CrCC;

            outputRow.AFF_BASE_VAR = CrBV;
            outputRow.AFF_MODALITY = CrModCode;
            outputRow.AFF_MKT_SEGMENT = CrMktSegment;
            outputRow.AFF_FOLDER = CrFolder;

            //判断AFF_FOLDER
            //if (LE.Length > 2)
            //{
            //    if (LE.Substring(0, 2).Equals("76"))
            //    {
            //        outputRow.AFF_FOLDER = "7081481";
            //    }
            //    else if (LE.Substring(0, 2).Equals("70"))
            //    {
            //        outputRow.AFF_FOLDER = "7081481";
            //    }
            //}
            //判断Source
            if (LE.Length > 2)
            {
                if (LE.Substring(0, 2).Equals("76"))
                {
                    outputRow.AFF_SOURCE = "7601";
                }
                else if (LE.Substring(0, 2).Equals("70"))
                {
                    outputRow.AFF_SOURCE = "7001";
                }
            }
            //判断AFF_DESTINATION
            if (LE.Length > 2)
            {
                if (LE.Substring(0, 2).Equals("76"))
                {
                    outputRow.AFF_DESTINATION = "76";
                }
                else if (LE.Substring(0, 2).Equals("70"))
                {
                    outputRow.AFF_DESTINATION = "70";
                }
            }
            decimal dMoney = 0;
            if (LE.Equals("760110"))
            {
                dMoney = Math.Round(iIns * dMor, 2);
            }
            else if (LE.Equals("700110"))
            {
                dMoney = iIns;
            }
            outputRow.ENTERED_DR = "";
            outputRow.ENTERED_CR = dMoney.ToString();
            outputRow.ACCOUNTED_DR = "";
            outputRow.ACCOUNTED_CR = "";
            //判断SET_OF_BOOKS_ID
            if (LE.Length > 2)
            {
                if (LE.Substring(0, 2).Equals("76"))
                {
                    outputRow.SET_OF_BOOKS_ID = "050";
                }
                else if (LE.Substring(0, 2).Equals("70"))
                {
                    outputRow.SET_OF_BOOKS_ID = "051";
                }
            }
            outputRow.ACTUAL_FLAG = "A";
            outputRow.AFF_JOURNAL_CATEGORY = "Adjustment";
            outputRow.AFF_JOURNAL_SOURCE = "Spreadsheet";
            outputRow.PERIOD = "";
            outputRow.CODE_COMBINATION_ID = "";
            outputRow.COMPANY_CODE_MAP = "";
            outputRow.JOURNAL_SOURCE_MAP = "";
            outputRow.JOURNAL_CATEGORY_MAP = "";
            outputRow.JOURNAL_BATCH = txtJOURNAL_BATCH.Text;
            outputRow.BATCH_DESCRIPTION = txtBATCH_DESCRIPTION.Text;
            outputRow.JOURNAL_NAME = txtJOURNAL_BATCH.Text;
            outputRow.JOURNAL_DESCRIPTION = txtBATCH_DESCRIPTION.Text;
            outputRow.JOURNAL_REFERENCE = "";
            outputRow.JOURNALLINEDESC = txtDESC_APPII.Text;
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
        /// 生成VCP数据
        /// </summary>
        /// <returns></returns>
        private int GenerateVCP()
        {
            //进行INS转换
            var DrAccount = "";
            var DrMktSegment = "";
            var DrFolder = "";
            var DrBV = "";
            var results = from rs in dsMain.ITEMMapping
                          where rs.Item == "VCP"
                          select rs; ;
            foreach (var atemp in results)
            {
                //DrAccount = atemp.DrAccount;
                DrMktSegment = atemp.DrMktSegment;
               // DrFolder = atemp.DrFolder;
                DrBV = atemp.DrBV;
                break;
            }

            var CrAccount = "";
            var CrMktSegment = "";
            var CrFolder = "";
            var CrBV = "";
            var CrResults = from rs in dsMain.ITEMMapping
                            where rs.Item == "VCP"
                            select rs; ;
            foreach (var atemp in CrResults)
            {
                CrAccount = atemp.CrAccount;
                CrMktSegment = atemp.CrMktSegment;
                CrFolder = atemp.CrGETDFolder;
                CrBV = atemp.CrBV;
                break;
            }

            pbMain.Value = 0;
            pbMain.Maximum = uGridInput.Rows.Count;
            var iCount = 0;
            for (var i = 0; i < uGridInput.Rows.Count; i++)
            {
                //获取INS数据
                decimal dwar;
                var cwar = uGridInput.Rows[i].Cells["VCP"].Value.ToString();
                var cModCode = uGridInput.Rows[i].Cells["MOD_Code"].Value.ToString();
                var bwar = decimal.TryParse(cwar, out dwar);
                if (!bwar)
                    continue;
                if (dwar == 0)
                    continue;
                //先写借
                WriteVCPDr(i, dwar, DrAccount, DrMktSegment, DrFolder, DrBV);


                pbMain.Value = i;
                iCount = iCount + 1;
            }

            //计算LE分类的合计
            iCount = iCount + VCPGroupByVCP(CrAccount, CrMktSegment, CrFolder, CrBV);


            return iCount;

        }
        /// <summary>
        /// 计算APPI的Groupby 结果
        /// </summary>
        /// <param name="CrAccount"></param>
        /// <param name="CrMktSegment"></param>
        /// <param name="CrFolder"></param>
        /// <param name="CrBV"></param>
        private int VCPGroupByVCP(string CrAccount, string CrMktSegment, string CrFolder, string CrBV)
        {
            var resultAPPI = from u in dsMain.DataInput
                             group u by new { LE = u.LE } into g
                             select new
                             {
                                 g.Key.LE,
                                 VCP = g.Sum(c => c.VCP)
                             };
            var iCount = 0;
            foreach (var itemAPPI in resultAPPI)
            {
                if (itemAPPI.VCP == 0)
                    return 0;
                WriteVCPCrGoupBy(itemAPPI.LE, "990", itemAPPI.VCP, CrAccount, CrMktSegment, CrFolder, CrBV);
                iCount = iCount + 1;
            }
            return iCount;
        }
        /// <summary>
        /// VCP输出写借方
        /// </summary>
        /// <param name="i"></param>
        private void WriteVCPDr(int i, decimal iIns, string DrAccount, string DrMktSegment, string DrFolder, string DrBV)
        {
            var citem = "VCP";
            var LE = uGridInput.Rows[i].Cells["LE"].Value.ToString();
            var PL = uGridInput.Rows[i].Cells["PL"].Value.ToString();
            var Mod = uGridInput.Rows[i].Cells["MOD_Code"].Value.ToString();
            var DrCC = "";
            var results = from rs in dsMain.SDMatrix
                          where rs.Item == "VCP" && rs.PL == PL && rs.LE == LE && rs.ModCode == Mod
                          select rs; ;

            foreach (var atemp in results)
            {
                DrAccount = atemp.DrAccount;
                DrCC = atemp.DrCC;
                DrFolder = atemp.DrFolder;
                break;
            }

            var outputRow = dsMain.DataOutPut.NewDataOutPutRow();
            outputRow.ACCOUNTING_DATE = txtACCOUNTING_DATE.Text;
            //判断币种
            var cCurrency = "";
            if (LE.Equals("760110"))
            {
                outputRow.CURRENCY_CODE = "CNY";
            }
            else if (LE.Equals("700110"))
            {
                outputRow.CURRENCY_CODE = "USD";
            }

            outputRow.DATE_CREATED = "";
            outputRow.CURRENCY_CONV_DATE = "";
            outputRow.CURRENCY_CONV_TYPE = "";
            outputRow.CURRENCY_CONV_RATE = "";
            //Company
            outputRow.AFF_COMPANY = LE;
            outputRow.AFF_ACCOUNT = DrAccount;
            //成本中心
            outputRow.AFF_CENTER = DrCC;

            outputRow.AFF_BASE_VAR = DrBV;
            outputRow.AFF_MODALITY = uGridInput.Rows[i].Cells["MOD_Code"].Value.ToString();
            outputRow.AFF_MKT_SEGMENT = DrMktSegment;
            outputRow.AFF_FOLDER = DrFolder;

            //判断AFF_FOLDER
            //if (LE.Length > 2)
            //{
            //    if (LE.Substring(0, 2).Equals("76"))
            //    {
            //        outputRow.AFF_FOLDER = "7081481";
            //    }
            //    else if (LE.Substring(0, 2).Equals("70"))
            //    {
            //        outputRow.AFF_FOLDER = "7081481";
            //    }
            //}
            //判断Source
            if (LE.Length > 2)
            {
                if (LE.Substring(0, 2).Equals("76"))
                {
                    outputRow.AFF_SOURCE = "7601";
                }
                else if (LE.Substring(0, 2).Equals("70"))
                {
                    outputRow.AFF_SOURCE = "7001";
                }
            }
            //判断AFF_DESTINATION
            outputRow.AFF_DESTINATION = "00";
            //if (LE.Length > 2)
            //{
            //    if (LE.Substring(0, 2).Equals("76"))
            //    {
            //        outputRow.AFF_DESTINATION = "76";
            //    }
            //    else if (LE.Substring(0, 2).Equals("70"))
            //    {
            //        outputRow.AFF_DESTINATION = "70";
            //    }
            //}
            decimal dMoney = 0;
            if (LE.Equals("760110"))
            {
                dMoney = Math.Round(iIns * dMor, 2);
            }
            else if (LE.Equals("700110"))
            {
                dMoney = iIns;
            }
            outputRow.ENTERED_DR = dMoney.ToString();
            outputRow.ENTERED_CR = "";
            outputRow.ACCOUNTED_DR = "";
            outputRow.ACCOUNTED_CR = "";
            //判断SET_OF_BOOKS_ID
            if (LE.Length > 2)
            {
                if (LE.Substring(0, 2).Equals("76"))
                {
                    outputRow.SET_OF_BOOKS_ID = "050";
                }
                else if (LE.Substring(0, 2).Equals("70"))
                {
                    outputRow.SET_OF_BOOKS_ID = "051";
                }
            }
            outputRow.ACTUAL_FLAG = "A";
            outputRow.AFF_JOURNAL_CATEGORY = "Adjustment";
            outputRow.AFF_JOURNAL_SOURCE = "Spreadsheet";
            outputRow.PERIOD = "";
            outputRow.CODE_COMBINATION_ID = "";
            outputRow.COMPANY_CODE_MAP = "";
            outputRow.JOURNAL_SOURCE_MAP = "";
            outputRow.JOURNAL_CATEGORY_MAP = "";
            outputRow.JOURNAL_BATCH = txtJOURNAL_BATCH.Text;
            outputRow.BATCH_DESCRIPTION = txtBATCH_DESCRIPTION.Text;
            outputRow.JOURNAL_NAME = txtJOURNAL_BATCH.Text;
            outputRow.JOURNAL_DESCRIPTION = txtBATCH_DESCRIPTION.Text;
            outputRow.JOURNAL_REFERENCE = "";
            outputRow.JOURNALLINEDESC = txtDESC_VCP.Text;
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
            outputRow.ORDER_NUM = uGridInput.Rows[i].Cells["OrderNumber"].Value.ToString();
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
        /// VCP输出写贷方
        /// </summary>
        /// <param name="i"></param>
        //private void WriteVCPCr(int i, decimal iIns, string CrAccount, string CrMktSegment, string CrFolder, string CrBV)
        //{
        //    var citem = "VCP";
        //    var LE = uGridInput.Rows[i].Cells["LE"].Value.ToString();
        //    var PL = uGridInput.Rows[i].Cells["PL"].Value.ToString();
        //    var Mod = uGridInput.Rows[i].Cells["MOD_Code"].Value.ToString();
        //    var CrCC = "";
        //    var CrModCode = "";
        //    var results = from rs in dsMain.SDMatrix
        //                  where rs.Item == "VCP" && rs.PL == PL && rs.LE == LE && rs.ModCode == Mod
        //                  select rs; ;

        //    foreach (var atemp in results)
        //    {
        //        CrAccount = atemp.CrAccount;
        //        CrCC = atemp.CrCC;
        //        CrFolder = atemp.CrFolder;
        //        CrModCode = atemp.CrModCode;
        //        break;
        //    }

        //    var outputRow = dsMain.DataOutPut.NewDataOutPutRow();
        //    outputRow.ACCOUNTING_DATE = txtACCOUNTING_DATE.Text;
        //    //判断币种
        //    var cCurrency = "";

        //    if (LE.Equals("760110"))
        //    {
        //        outputRow.CURRENCY_CODE = "CNY";
        //    }
        //    else if (LE.Equals("700110"))
        //    {
        //        outputRow.CURRENCY_CODE = "USD";
        //    }

        //    outputRow.DATE_CREATED = "";
        //    outputRow.CURRENCY_CONV_DATE = "";
        //    outputRow.CURRENCY_CONV_TYPE = "";
        //    outputRow.CURRENCY_CONV_RATE = "";
        //    //Company
        //    outputRow.AFF_COMPANY = LE;
        //    outputRow.AFF_ACCOUNT = CrAccount;
        //    //成本中心

        //    outputRow.AFF_CENTER = CrCC;

        //    outputRow.AFF_BASE_VAR = CrBV;
        //    outputRow.AFF_MODALITY = CrModCode;
        //    outputRow.AFF_MKT_SEGMENT = CrMktSegment;
        //    outputRow.AFF_FOLDER = CrFolder;

        //    //判断AFF_FOLDER
        //    //if (LE.Length > 2)
        //    //{
        //    //    if (LE.Substring(0, 2).Equals("76"))
        //    //    {
        //    //        outputRow.AFF_FOLDER = "7081481";
        //    //    }
        //    //    else if (LE.Substring(0, 2).Equals("70"))
        //    //    {
        //    //        outputRow.AFF_FOLDER = "7081481";
        //    //    }
        //    //}
        //    //判断Source
        //    if (LE.Length > 2)
        //    {
        //        if (LE.Substring(0, 2).Equals("76"))
        //        {
        //            outputRow.AFF_SOURCE = "7601";
        //        }
        //        else if (LE.Substring(0, 2).Equals("70"))
        //        {
        //            outputRow.AFF_SOURCE = "7001";
        //        }
        //    }
        //    //判断AFF_DESTINATION
        //    if (LE.Length > 2)
        //    {
        //        if (LE.Substring(0, 2).Equals("76"))
        //        {
        //            outputRow.AFF_DESTINATION = "76";
        //        }
        //        else if (LE.Substring(0, 2).Equals("70"))
        //        {
        //            outputRow.AFF_DESTINATION = "70";
        //        }
        //    }
        //    decimal dMoney = 0;
        //    if (LE.Equals("760110"))
        //    {
        //        dMoney = Math.Round(iIns * dMor, 2);
        //    }
        //    else if (LE.Equals("700110"))
        //    {
        //        dMoney = iIns;
        //    }
        //    outputRow.ENTERED_DR = "";
        //    outputRow.ENTERED_CR = dMoney.ToString();
        //    outputRow.ACCOUNTED_DR = "";
        //    outputRow.ACCOUNTED_CR = "";
        //    //判断SET_OF_BOOKS_ID
        //    if (LE.Length > 2)
        //    {
        //        if (LE.Substring(0, 2).Equals("76"))
        //        {
        //            outputRow.SET_OF_BOOKS_ID = "050";
        //        }
        //        else if (LE.Substring(0, 2).Equals("70"))
        //        {
        //            outputRow.SET_OF_BOOKS_ID = "051";
        //        }
        //    }
        //    outputRow.ACTUAL_FLAG = "A";
        //    outputRow.AFF_JOURNAL_CATEGORY = "Adjustment";
        //    outputRow.AFF_JOURNAL_SOURCE = "Spreadsheet";
        //    outputRow.PERIOD = "";
        //    outputRow.CODE_COMBINATION_ID = "";
        //    outputRow.COMPANY_CODE_MAP = "";
        //    outputRow.JOURNAL_SOURCE_MAP = "";
        //    outputRow.JOURNAL_CATEGORY_MAP = "";
        //    outputRow.JOURNAL_BATCH = txtJOURNAL_BATCH.Text;
        //    outputRow.BATCH_DESCRIPTION = txtBATCH_DESCRIPTION.Text;
        //    outputRow.JOURNAL_NAME = txtJOURNAL_BATCH.Text;
        //    outputRow.JOURNAL_DESCRIPTION = txtBATCH_DESCRIPTION.Text;
        //    outputRow.JOURNAL_REFERENCE = "";
        //    outputRow.JOURNALLINEDESC = txtDESC_VCP.Text;
        //    outputRow.LEGACY_ACCOUNT = "";
        //    outputRow.LEGACY_JRNL_NUM = "";
        //    outputRow.LEGACY_OFFSET_ACCT = "";
        //    outputRow.BILL_TO_CUSTOMER = "";
        //    outputRow.SHIP_TO_CUSTOMER = "";
        //    outputRow.EMPLOYEE_NUM = "";
        //    outputRow.INVENTORY_ORG = "";
        //    outputRow.MA_CODE = "";
        //    outputRow.MATERIAL_CLASS = "";
        //    outputRow.PO_ITEM = "";
        //    outputRow.ORDER_NUM = uGridInput.Rows[i].Cells["OrderNumber"].Value.ToString();
        //    outputRow.INV_ITEM_NUM = "";
        //    outputRow.QUANTITY = "";
        //    outputRow.UNIT_OF_MEASURE = "";
        //    outputRow.DOCUMENT_NUM = "";
        //    outputRow.DOCUMENT_DATE = "";
        //    outputRow.PROJECT_NUM = "";
        //    outputRow.DOCUMENT_NUM2 = "";
        //    outputRow.SHIPPED_DATE = "";
        //    outputRow.VAT_CODE = "";
        //    outputRow.ACTUAL_HOURS = "";
        //    outputRow.CONSIGNMT_CONTRACT = "";
        //    outputRow.COST_KEY = "";
        //    outputRow.PO_NUM = "";
        //    outputRow.PSI_CODE = "";
        //    outputRow.RETURN_MAT_CODE = "";
        //    outputRow.TRANSACTION_CODE = "";
        //    outputRow.VENDOR_NUM = "";
        //    outputRow.SERVICE_ACCTG_KEY = "";
        //    outputRow.REFERENCE_AMOUNT = "";
        //    outputRow.LOCAL_MAPPING_FIELD1 = "";
        //    outputRow.LOCAL_MAPPING_FIELD2 = "";
        //    outputRow.LOCAL_MAPPING_FIELD3 = "";
        //    outputRow.LOCAL_MAPPING_FIELD4 = "";
        //    outputRow.LOCAL_MAPPING_FIELD5 = "";
        //    outputRow.LOCAL_MAPPING_FIELD6 = "";
        //    outputRow.LOCAL_MAPPING_FIELD7 = "";
        //    outputRow.LOCAL_MAPPING_FIELD8 = "";
        //    outputRow.LOCAL_MAPPING_FIELD9 = "";
        //    outputRow.LOCAL_MAPPING_FIELD10 = "";
        //    outputRow.LOCAL_MAPPING_FIELD11 = "";
        //    outputRow.LOCAL_MAPPING_FIELD12 = "";
        //    outputRow.LOCAL_MAPPING_FIELD13 = "";
        //    outputRow.LOCAL_MAPPING_FIELD14 = "";
        //    outputRow.LOCAL_MAPPING_FIELD15 = "";
        //    outputRow.LOCAL_MAPPING_FIELD16 = "";
        //    outputRow.LOCAL_MAPPING_FIELD17 = "";
        //    outputRow.LOCAL_MAPPING_FIELD18 = "";
        //    outputRow.LOCAL_MAPPING_FIELD19 = "";
        //    dsMain.DataOutPut.Rows.Add(outputRow);
        //}
        /// <summary>
        /// 写VCPGroupby后的贷方数据
        /// </summary>
        /// <param name="LE"></param>
        /// <param name="ModCode"></param>
        /// <param name="iIns"></param>
        /// <param name="CrAccount"></param>
        /// <param name="CrMktSegment"></param>
        /// <param name="CrFolder"></param>
        /// <param name="CrBV"></param>
        private void WriteVCPCrGoupBy(string LE, string ModCode, decimal iIns, string CrAccount, string CrMktSegment, string CrFolder, string CrBV)
        {

            //var citem = "VCP";
            //var PL ="L23";

            //var CrCC = "";
            //var CrModCode = "";
            //var results = from rs in dsMain.SDMatrix
            //              where rs.Item == "VCP" && rs.PL == PL && rs.LE == LE && rs.ModCode == ModCode
            //              select rs; ;

            //foreach (var atemp in results)
            //{
            //    CrAccount = atemp.CrAccount;
            //    CrCC = atemp.CrCC;
            //    CrFolder = atemp.CrFolder;
            //    CrModCode = atemp.CrModCode;
            //    break;
            //}
            var outputRow = dsMain.DataOutPut.NewDataOutPutRow();
            outputRow.ACCOUNTING_DATE = txtACCOUNTING_DATE.Text;
            //判断币种
            var cCurrency = "";
            if (LE.Equals("760110"))
            {
                outputRow.CURRENCY_CODE = "CNY";
            }
            else if (LE.Equals("700110"))
            {
                outputRow.CURRENCY_CODE = "USD";
            }

            outputRow.DATE_CREATED = "";
            outputRow.CURRENCY_CONV_DATE = "";
            outputRow.CURRENCY_CONV_TYPE = "";
            outputRow.CURRENCY_CONV_RATE = "";
            //Company
            outputRow.AFF_COMPANY = LE;
            outputRow.AFF_ACCOUNT = CrAccount;
            //成本中心
            if (LE.Length > 2)
            {
                if (LE.Substring(0, 2).Equals("76"))
                {
                    outputRow.AFF_CENTER = "761625";
                }
                else if (LE.Substring(0, 2).Equals("70"))
                {
                    outputRow.AFF_CENTER = "701625";
                }
            }

            outputRow.AFF_BASE_VAR = CrBV;
            //VCP的Cr. MOD code should all be 990
            outputRow.AFF_MODALITY = ModCode;
            outputRow.AFF_MKT_SEGMENT = CrMktSegment;
            outputRow.AFF_FOLDER = CrFolder;

            //判断AFF_FOLDER
            //if (LE.Length > 2)
            //{
            //    if (LE.Substring(0, 2).Equals("76"))
            //    {
            //        outputRow.AFF_FOLDER = "7081481";
            //    }
            //    else if (LE.Substring(0, 2).Equals("70"))
            //    {
            //        outputRow.AFF_FOLDER = "7081481";
            //    }
            //}
            //判断Source
            if (LE.Length > 2)
            {
                if (LE.Substring(0, 2).Equals("76"))
                {
                    outputRow.AFF_SOURCE = "7601";
                }
                else if (LE.Substring(0, 2).Equals("70"))
                {
                    outputRow.AFF_SOURCE = "7001";
                }
            }
            //判断AFF_DESTINATION
            outputRow.AFF_DESTINATION = "00";
            //if (LE.Length > 2)
            //{
            //    if (LE.Substring(0, 2).Equals("76"))
            //    {
            //        outputRow.AFF_DESTINATION = "76";
            //    }
            //    else if (LE.Substring(0, 2).Equals("70"))
            //    {
            //        outputRow.AFF_DESTINATION = "70";
            //    }
            //}
            decimal dMoney = 0;
            if (LE.Equals("760110"))
            {
                dMoney = Math.Round(iIns * dMor, 2);
            }
            else if (LE.Equals("700110"))
            {
                dMoney = iIns;
            }
            outputRow.ENTERED_DR = "";
            outputRow.ENTERED_CR = dMoney.ToString();
            outputRow.ACCOUNTED_DR = "";
            outputRow.ACCOUNTED_CR = "";
            //判断SET_OF_BOOKS_ID
            if (LE.Length > 2)
            {
                if (LE.Substring(0, 2).Equals("76"))
                {
                    outputRow.SET_OF_BOOKS_ID = "050";
                }
                else if (LE.Substring(0, 2).Equals("70"))
                {
                    outputRow.SET_OF_BOOKS_ID = "051";
                }
            }
            outputRow.ACTUAL_FLAG = "A";
            outputRow.AFF_JOURNAL_CATEGORY = "Adjustment";
            outputRow.AFF_JOURNAL_SOURCE = "Spreadsheet";
            outputRow.PERIOD = "";
            outputRow.CODE_COMBINATION_ID = "";
            outputRow.COMPANY_CODE_MAP = "";
            outputRow.JOURNAL_SOURCE_MAP = "";
            outputRow.JOURNAL_CATEGORY_MAP = "";
            outputRow.JOURNAL_BATCH = txtJOURNAL_BATCH.Text;
            outputRow.BATCH_DESCRIPTION = txtBATCH_DESCRIPTION.Text;
            outputRow.JOURNAL_NAME = txtJOURNAL_BATCH.Text;
            outputRow.JOURNAL_DESCRIPTION = txtBATCH_DESCRIPTION.Text;
            outputRow.JOURNAL_REFERENCE = "";
            outputRow.JOURNALLINEDESC = txtDESC_VCP.Text;
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
        /// 导出Excel表格数据
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void tsbtnExportExcel_Click(object sender, EventArgs e)
        {
            sfdMain.Filter = @"Excel2003文件(*.xls)|*.xls";
            sfdMain.DefaultExt = "xls";
            sfdMain.Title = @"文件保存到";
            sfdMain.FileName = DateTime.Today.ToString("yyyyMMdd");
            if (sfdMain.ShowDialog() != DialogResult.OK) return;
            try
            {
                //uGridExport2007.Export(uGridOutput, sfdMain.FileName);
                //MessageBox.Show(@"Export Success，File Path：" + sfdMain.FileName, @"Success");
                //tslblProgress.Text = "Export Success";
                OutFileToDisk(dsMain.DataOutPut, "Output", sfdMain.FileName);
                MessageBox.Show(@"Export Success，File Path：" + sfdMain.FileName, @"Success");
                tslblProgress.Text = "Export Success";
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }

        }

        //private void tsbtnExportCsv_Click(object sender, EventArgs e)
        //{
        //    ExportDataGridToCSV(dsMain.DataOutPut);
        //}

        ///// <summary>
        ///// Export the data from datatable to CSV file
        ///// </summary>
        ///// <param name="grid"></param>
        //public void ExportDataGridToCSV(DataTable dt)
        //{
        //    //File info initialization
        //    sfdMain.Filter = @"CSV文件(*.csv)|*.csv";
        //    sfdMain.DefaultExt = "csv";
        //    sfdMain.Title = @"文件保存到";
        //    sfdMain.FileName = DateTime.Today.ToString("yyyyMMdd");
        //    if (sfdMain.ShowDialog() != DialogResult.OK) return;

        //    System.IO.FileStream fs = new FileStream(sfdMain.FileName, System.IO.FileMode.Create, System.IO.FileAccess.Write);
        //    StreamWriter sw = new StreamWriter(fs, new System.Text.UnicodeEncoding());
        //    //Tabel header
        //    for (int i = 0; i < dt.Columns.Count; i++)
        //    {
        //        sw.Write(dt.Columns[i].ColumnName);
        //        sw.Write("\t");
        //    }
        //    sw.WriteLine("");
        //    //Table body
        //    for (int i = 0; i < dt.Rows.Count; i++)
        //    {
        //        for (int j = 0; j < dt.Columns.Count; j++)
        //        {
        //            sw.Write(dt.Rows[i][j].ToString());
        //            sw.Write("\t");
        //        }
        //        sw.WriteLine("");
        //    }
        //    sw.Flush();
        //    sw.Close();
        //    MessageBox.Show("Save CSV:" + sfdMain.FileName);
        //}

        private void uGridExport2007_BeginExport(object sender, Infragistics.Win.UltraWinGrid.ExcelExport.BeginExportEventArgs e)
        {
            pbMain.Value = 0;
            pbMain.Maximum = uGridOutput.Rows.Count;

        }

        private void uGridExport2007_RowExported(object sender, Infragistics.Win.UltraWinGrid.ExcelExport.RowExportedEventArgs e)
        {
            pbMain.Value = e.CurrentRowIndex;
        }

        /// <summary> 
        /// 导出数据到本地 Excel
        /// </summary> 
        /// <param name="dt">要导出的数据</param> 
        /// <param name="tableName">表格标题</param> 
        /// <param name="path">保存路径</param> 
        public static void OutFileToDisk(DataTable dt, string tableName, string path)
        {


            Aspose.Cells.Workbook workbook = new Aspose.Cells.Workbook(); //工作簿 
            Aspose.Cells.Worksheet sheet = workbook.Worksheets[0]; //工作表 
            Cells cells = sheet.Cells;//单元格 

            //为标题设置样式     
            //Style styleTitle = workbook.Styles[workbook.Styles.Add()];//新增样式 
            //styleTitle.HorizontalAlignment = TextAlignmentType.Center;//文字居中 
            //styleTitle.Font.Name = "宋体";//文字字体 
            //styleTitle.Font.Size = 18;//文字大小 
            //styleTitle.Font.IsBold = true;//粗体 

            ////样式2 
            //Style style2 = workbook.Styles[workbook.Styles.Add()];//新增样式 
            //style2.HorizontalAlignment = TextAlignmentType.Center;//文字居中 
            //style2.Font.Name = "宋体";//文字字体 
            //style2.Font.Size = 14;//文字大小 
            //style2.Font.IsBold = true;//粗体 
            //style2.IsTextWrapped = true;//单元格内容自动换行 
            //style2.Borders[BorderType.LeftBorder].LineStyle = CellBorderType.Thin;
            //style2.Borders[BorderType.RightBorder].LineStyle = CellBorderType.Thin;
            //style2.Borders[BorderType.TopBorder].LineStyle = CellBorderType.Thin;
            //style2.Borders[BorderType.BottomBorder].LineStyle = CellBorderType.Thin;

            ////样式3 
            //Style style3 = workbook.Styles[workbook.Styles.Add()];//新增样式 
            //style3.HorizontalAlignment = TextAlignmentType.Center;//文字居中 
            //style3.Font.Name = "宋体";//文字字体 
            //style3.Font.Size = 12;//文字大小 
            //style3.Borders[BorderType.LeftBorder].LineStyle = CellBorderType.Thin;
            //style3.Borders[BorderType.RightBorder].LineStyle = CellBorderType.Thin;
            //style3.Borders[BorderType.TopBorder].LineStyle = CellBorderType.Thin;
            //style3.Borders[BorderType.BottomBorder].LineStyle = CellBorderType.Thin;

            int Colnum = dt.Columns.Count;//表格列数 
            int Rownum = dt.Rows.Count;//表格行数 


            //生成行2 列名行 
            for (int i = 0; i < Colnum; i++)
            {
                cells[0, i].PutValue(dt.Columns[i].ColumnName);
                cells.SetColumnWidth(i, 25);
                //cells[0, i].SetStyle(style2);
            }

            //生成数据行 
            for (int i = 0; i < Rownum; i++)
            {
                for (int k = 0; k < Colnum; k++)
                {
                    //if(cells[1 + i].Name.Equals("ENTERED_DR")||cells[1 + i].Name.Equals("ENTERED_CR"))
                    //{
                    //    if (string.IsNullOrEmpty(dt.Rows[i][k])
                    //    cells[1 + i, k].PutValue(dt.Rows[i][k]);
                    //}
                    cells[1 + i, k].PutValue(dt.Rows[i][k]);
                    //cells[1 + i, k].SetStyle(style3);
                }
            }

            workbook.Save(path);
        }

    }
}
