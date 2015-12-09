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
                   
                    inputRow.SetColumnError("INS", "SouceData: "+cins + "    cannot convert to decimal");
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
                    inputRow.SetColumnError("APPI", "SouceData: "+cappi + "  cannot convert to decimal");
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
            for (var i = 0; i < dsMain.DataInput.Rows.Count ; i++)
            {
                decimal dins,dwar,dappi,dappii,dvcp;
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
            if(dsMain.DataInput.HasErrors)
            {
                tslblProgress.Text = @"Input DataSouce Find Some Error, Cannot covert some data to Decimal Type";
            }
        }

        private void tsbtnTransferData_Click(object sender, EventArgs e)
        {
            dsMain.DataOutPut.Rows.Clear();
            tcMain.SelectedTab=tcMain.TabPages[1];
            var iInsCount = GenerateINS(); 
            var iWarCount = GenerateWAR();


            //绘颜色
            for (var i = 0; i < iInsCount; i++)
            {
                uGridOutput.Rows[i].Appearance.BackColor = Color.LightGreen;
                uGridOutput.Rows[i].Activate();
                Application.DoEvents();
            }

            for (var i = iInsCount; i < iInsCount + iWarCount; i++)
            {
                uGridOutput.Rows[i].Appearance.BackColor = Color.LightGreen;
                uGridOutput.Rows[i].Activate();
                Application.DoEvents();
            }
        }


        private int GenerateINS()
        {
            //进行INS转换
            var DrAccount = "561000013";
            var DrMktSegment = "90";
            var DrFolder = "0000000";
            var DrBV = "1";
            var results = from rs in dsMain.ITEMMapping
                          where rs.Item == "Installation"
                          select rs; ;
            foreach (var atemp in results)
            {
                DrAccount = atemp.DrAccount;
                DrMktSegment = atemp.DrMktSegment;
                DrFolder = atemp.DrFolder;
                DrBV = atemp.DrBV;
                break;
            }

            var CrAccount = "391030010";
            var CrMktSegment = "00";
            var CrFolder = "7681481";
            var CrBV = "0";
            var CrResults = from rs in dsMain.ITEMMapping
                          where rs.Item == "Installation"
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
            var iCount=0;
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
                WriteINSDr(i, dins, DrAccount, DrMktSegment, DrFolder, DrBV);
                WriteINSCr(i, dins, CrAccount, CrMktSegment, CrFolder, CrBV);
                pbMain.Value = i;
                iCount = iCount + 1;
            }
            return iCount;
            
        }
        /// <summary>
        /// INS输出写借方
        /// </summary>
        /// <param name="i"></param>
        private void WriteINSDr(int i, decimal iIns, string DrAccount, string DrMktSegment, string DrFolder,string DrBV)
        {
            var outputRow = dsMain.DataOutPut.NewDataOutPutRow();
            outputRow.ACCOUNTING_DATE = txtACCOUNTING_DATE.Text;
            //判断币种
            var cCurrency = "";
            var LE = uGridInput.Rows[i].Cells["LE"].Value.ToString();
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
            outputRow.AFF_COMPANY =LE;
            outputRow.AFF_ACCOUNT = DrAccount;
            //成本中心
            var PL = uGridInput.Rows[i].Cells["PL"].Value.ToString();
            var results = from rs in dsMain.LEMapping
                          where rs.PL == PL&&rs.LE==LE
                          select rs; ;
            foreach (var atemp in results)
            {
                outputRow.AFF_CENTER = atemp.CC;
                break;
            }

            outputRow.AFF_BASE_VAR = DrBV;
            outputRow.AFF_MODALITY = uGridInput.Rows[i].Cells["MOD_Code"].Value.ToString();
            outputRow.AFF_MKT_SEGMENT = DrMktSegment;
            outputRow.AFF_FOLDER = DrFolder;

            //判断AFF_FOLDER
            if (LE.Length > 2)
            {
                if (LE.Substring(0, 2).Equals("76"))
                {
                    outputRow.AFF_FOLDER = "7081481";
                }
                else if (LE.Substring(0, 2).Equals("70"))
                {
                    outputRow.AFF_FOLDER = "7081481";
                }
            }
            //判断Source
            if(LE.Length>2)
            {
                if(LE.Substring(0,2).Equals("76"))
                {
                    outputRow.AFF_SOURCE = "7601";
                }
                else if(LE.Substring(0,2).Equals("70"))
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

            outputRow.ENTERED_DR = iIns.ToString() ;
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
            outputRow.JOURNALLINEDESC = txtJOURNALLINEDESC.Text;
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
            outputRow.ORDER_NUM = uGridInput.Rows[i].Cells["OrderNumber"].ToString();
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
        private void WriteINSCr(int i,decimal iIns,string CrAccount, string CrMktSegment, string CrFolder,string CrBV)
        {
            var outputRow = dsMain.DataOutPut.NewDataOutPutRow();
            outputRow.ACCOUNTING_DATE = txtACCOUNTING_DATE.Text;
            //判断币种
            var cCurrency = "";
            var LE = uGridInput.Rows[i].Cells["LE"].Value.ToString();
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
            var PL = uGridInput.Rows[i].Cells["PL"].Value.ToString();
            var results = from rs in dsMain.LEMapping
                          where rs.PL == PL && rs.LE == LE
                          select rs; ;
            foreach (var atemp in results)
            {
                outputRow.AFF_CENTER = atemp.CC;
                break;
            }

            outputRow.AFF_BASE_VAR = CrBV;
            outputRow.AFF_MODALITY = uGridInput.Rows[i].Cells["MOD_Code"].Value.ToString();
            outputRow.AFF_MKT_SEGMENT = CrMktSegment;
            outputRow.AFF_FOLDER = CrFolder;

            //判断AFF_FOLDER
            if (LE.Length > 2)
            {
                if (LE.Substring(0, 2).Equals("76"))
                {
                    outputRow.AFF_FOLDER = "7081481";
                }
                else if (LE.Substring(0, 2).Equals("70"))
                {
                    outputRow.AFF_FOLDER = "7081481";
                }
            }
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

            outputRow.ENTERED_DR = "";
            outputRow.ENTERED_CR = iIns.ToString();
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
            outputRow.JOURNALLINEDESC = txtJOURNALLINEDESC.Text;
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
            outputRow.ORDER_NUM = uGridInput.Rows[i].Cells["OrderNumber"].ToString();
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



        private int GenerateWAR()
        {
            //进行INS转换
            var DrAccount = "562000013";
            var DrMktSegment = "90";
            var DrFolder = "0000000";
            var DrBV = "1";
            var results = from rs in dsMain.ITEMMapping
                          where rs.Item == "Warranty"
                          select rs; ;
            foreach (var atemp in results)
            {
                DrAccount = atemp.DrAccount;
                DrMktSegment = atemp.DrMktSegment;
                DrFolder = atemp.DrFolder;
                DrBV = atemp.DrBV;
                break;
            }

            var CrAccount = "421001020";
            var CrMktSegment = "00";
            var CrFolder = "7082151";
            var CrBV = "0";
            var CrResults = from rs in dsMain.ITEMMapping
                            where rs.Item == "Warranty"
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
                var cwar = uGridInput.Rows[i].Cells["WAR"].Value.ToString();
                var bwar = decimal.TryParse(cwar, out dwar);
                if (!bwar)
                    continue;
                if (dwar == 0)
                    continue;
                //先写借
                WriteWARDr(i, dwar, DrAccount, DrMktSegment, DrFolder, DrBV);
                WriteWARCr(i, dwar, CrAccount, CrMktSegment, CrFolder, CrBV);
                pbMain.Value = i;
                iCount = iCount + 1;
            }
            return iCount;

        }
        /// <summary>
        /// WAR输出写借方
        /// </summary>
        /// <param name="i"></param>
        private void WriteWARDr(int i, decimal iIns, string DrAccount, string DrMktSegment, string DrFolder, string DrBV)
        {
            var outputRow = dsMain.DataOutPut.NewDataOutPutRow();
            outputRow.ACCOUNTING_DATE = txtACCOUNTING_DATE.Text;
            //判断币种
            var cCurrency = "";
            var LE = uGridInput.Rows[i].Cells["LE"].Value.ToString();
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
            var PL = uGridInput.Rows[i].Cells["PL"].Value.ToString();
            var results = from rs in dsMain.LEMapping
                          where rs.PL == PL && rs.LE == LE
                          select rs; ;
            foreach (var atemp in results)
            {
                outputRow.AFF_CENTER = atemp.CC;
                break;
            }

            outputRow.AFF_BASE_VAR = DrBV;
            outputRow.AFF_MODALITY = uGridInput.Rows[i].Cells["MOD_Code"].Value.ToString();
            outputRow.AFF_MKT_SEGMENT = DrMktSegment;
            outputRow.AFF_FOLDER = DrFolder;

            //判断AFF_FOLDER
            if (LE.Length > 2)
            {
                if (LE.Substring(0, 2).Equals("76"))
                {
                    outputRow.AFF_FOLDER = "7081481";
                }
                else if (LE.Substring(0, 2).Equals("70"))
                {
                    outputRow.AFF_FOLDER = "7081481";
                }
            }
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

            outputRow.ENTERED_DR = iIns.ToString();
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
            outputRow.JOURNALLINEDESC = txtJOURNALLINEDESC.Text;
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
            outputRow.ORDER_NUM = uGridInput.Rows[i].Cells["OrderNumber"].ToString();
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
        private void WriteWARCr(int i, decimal iIns, string CrAccount, string CrMktSegment, string CrFolder, string CrBV)
        {
            var outputRow = dsMain.DataOutPut.NewDataOutPutRow();
            outputRow.ACCOUNTING_DATE = txtACCOUNTING_DATE.Text;
            //判断币种
            var cCurrency = "";
            var LE = uGridInput.Rows[i].Cells["LE"].Value.ToString();
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
            var PL = uGridInput.Rows[i].Cells["PL"].Value.ToString();
            var results = from rs in dsMain.LEMapping
                          where rs.PL == PL && rs.LE == LE
                          select rs; ;
            foreach (var atemp in results)
            {
                outputRow.AFF_CENTER = atemp.CC;
                break;
            }

            outputRow.AFF_BASE_VAR = CrBV;
            outputRow.AFF_MODALITY = uGridInput.Rows[i].Cells["MOD_Code"].Value.ToString();
            outputRow.AFF_MKT_SEGMENT = CrMktSegment;
            outputRow.AFF_FOLDER = CrFolder;

            //判断AFF_FOLDER
            if (LE.Length > 2)
            {
                if (LE.Substring(0, 2).Equals("76"))
                {
                    outputRow.AFF_FOLDER = "7081481";
                }
                else if (LE.Substring(0, 2).Equals("70"))
                {
                    outputRow.AFF_FOLDER = "7081481";
                }
            }
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

            outputRow.ENTERED_DR = "";
            outputRow.ENTERED_CR = iIns.ToString();
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
            outputRow.JOURNALLINEDESC = txtJOURNALLINEDESC.Text;
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
            outputRow.ORDER_NUM = uGridInput.Rows[i].Cells["OrderNumber"].ToString();
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




        private int GenerateAPPI()
        {
            //进行INS转换
            var DrAccount = "562000013";
            var DrMktSegment = "90";
            var DrFolder = "0000000";
            var DrBV = "1";
            var results = from rs in dsMain.ITEMMapping
                          where rs.Item == "Warranty"
                          select rs; ;
            foreach (var atemp in results)
            {
                DrAccount = atemp.DrAccount;
                DrMktSegment = atemp.DrMktSegment;
                DrFolder = atemp.DrFolder;
                DrBV = atemp.DrBV;
                break;
            }

            var CrAccount = "421001020";
            var CrMktSegment = "00";
            var CrFolder = "7082151";
            var CrBV = "0";
            var CrResults = from rs in dsMain.ITEMMapping
                            where rs.Item == "Warranty"
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
                }
                
                pbMain.Value = i;
                iCount = iCount + 1;
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
                }
                pbMain.Value = i;
                iCount = iCount + 1;
            }

            WriteAPPICr(i, dwar, CrAccount, CrMktSegment, CrFolder, CrBV);
            
            return iCount;

        }
        /// <summary>
        /// WAR输出写借方
        /// </summary>
        /// <param name="i"></param>
        private void WriteAPPIDr(int i, decimal iIns, string DrAccount, string DrMktSegment, string DrFolder, string DrBV)
        {
            var outputRow = dsMain.DataOutPut.NewDataOutPutRow();
            outputRow.ACCOUNTING_DATE = txtACCOUNTING_DATE.Text;
            //判断币种
            var cCurrency = "";
            var LE = uGridInput.Rows[i].Cells["LE"].Value.ToString();
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
            var PL = uGridInput.Rows[i].Cells["PL"].Value.ToString();
            var results = from rs in dsMain.LEMapping
                          where rs.PL == PL && rs.LE == LE
                          select rs; ;
            foreach (var atemp in results)
            {
                outputRow.AFF_CENTER = atemp.CC;
                break;
            }

            outputRow.AFF_BASE_VAR = DrBV;
            outputRow.AFF_MODALITY = uGridInput.Rows[i].Cells["MOD_Code"].Value.ToString();
            outputRow.AFF_MKT_SEGMENT = DrMktSegment;
            outputRow.AFF_FOLDER = DrFolder;

            //判断AFF_FOLDER
            if (LE.Length > 2)
            {
                if (LE.Substring(0, 2).Equals("76"))
                {
                    outputRow.AFF_FOLDER = "7081481";
                }
                else if (LE.Substring(0, 2).Equals("70"))
                {
                    outputRow.AFF_FOLDER = "7081481";
                }
            }
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

            outputRow.ENTERED_DR = iIns.ToString();
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
            outputRow.JOURNALLINEDESC = txtJOURNALLINEDESC.Text;
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
            outputRow.ORDER_NUM = uGridInput.Rows[i].Cells["OrderNumber"].ToString();
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
        private void WriteAPPICr(int i, decimal iIns, string CrAccount, string CrMktSegment, string CrFolder, string CrBV)
        {
            var outputRow = dsMain.DataOutPut.NewDataOutPutRow();
            outputRow.ACCOUNTING_DATE = txtACCOUNTING_DATE.Text;
            //判断币种
            var cCurrency = "";
            var LE = uGridInput.Rows[i].Cells["LE"].Value.ToString();
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
            var PL = uGridInput.Rows[i].Cells["PL"].Value.ToString();
            var results = from rs in dsMain.LEMapping
                          where rs.PL == PL && rs.LE == LE
                          select rs; ;
            foreach (var atemp in results)
            {
                outputRow.AFF_CENTER = atemp.CC;
                break;
            }

            outputRow.AFF_BASE_VAR = CrBV;
            outputRow.AFF_MODALITY = uGridInput.Rows[i].Cells["MOD_Code"].Value.ToString();
            outputRow.AFF_MKT_SEGMENT = CrMktSegment;
            outputRow.AFF_FOLDER = CrFolder;

            //判断AFF_FOLDER
            if (LE.Length > 2)
            {
                if (LE.Substring(0, 2).Equals("76"))
                {
                    outputRow.AFF_FOLDER = "7081481";
                }
                else if (LE.Substring(0, 2).Equals("70"))
                {
                    outputRow.AFF_FOLDER = "7081481";
                }
            }
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

            outputRow.ENTERED_DR = "";
            outputRow.ENTERED_CR = iIns.ToString();
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
            outputRow.JOURNALLINEDESC = txtJOURNALLINEDESC.Text;
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
            outputRow.ORDER_NUM = uGridInput.Rows[i].Cells["OrderNumber"].ToString();
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
