using SAPbobsCOM;
using SAPbouiCOM;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Subtitution.Processor
{
    public class UpdateBomVer
    {
        private SAPbouiCOM.Application oSBOApplication;
        private SAPbobsCOM.Company oSBOCompany;

        /// <summary>
        /// Constructor
        /// </summary>
        public UpdateBomVer(SAPbouiCOM.Application oSBOApplication, SAPbobsCOM.Company oSBOCompany)
        {
            this.oSBOApplication = oSBOApplication;
            this.oSBOCompany = oSBOCompany;
        }

        /// <summary>
        /// Menu Event Update BOM Version
        /// When click menu, this event called
        /// </summary>
        public void MenuEvent_UpdateBomVer(ref MenuEvent pVal, out bool bubbleEvent)
        {
            bubbleEvent = true;
            if (pVal.BeforeAction == false)
            {
                Form oForm = null;

                try
                {
                    oForm = Utils.createForm(ref oSBOApplication, "Update BOM Version");
                    TemplateLoad(ref oForm);
                    oForm.Visible = true;
                }
                catch (Exception ex)
                {
                    bubbleEvent = false;
                    oSBOApplication.StatusBar.SetText(ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                }
                finally
                {
                    if (oForm != null)
                    {
                        if (bubbleEvent)
                        {
                            oForm.Freeze(false);
                            oForm.VisibleEx = true;
                        }
                        else
                            oForm.Close();
                    }
                    Utils.releaseObject(oForm);
                }
            }
        }

        private void TemplateLoad(ref Form oForm)
        {
            try
            {
                DBDataSource dtSource = null;
                dtSource = oForm.DataSources.DBDataSources.Item("@SOL_UPBOMVER_H");
                oForm.Freeze(true);

                int runCode = GenerateCode();
                string updDate = DateTime.Now.Date.ToShortDateString();
                string updTime = DateTime.Now.TimeOfDay.ToString();
                GetLastUpdate(out updDate, out updTime);

                oForm.Items.Item("tUpDt").Enabled = true;
                oForm.Items.Item("tUpTm").Enabled = true;

                dtSource.SetValue("Code", 0, runCode.ToString());
                dtSource.SetValue("U_SOL_UPDATE", 0, updDate);
                dtSource.SetValue("U_SOL_UPTIME", 0, updTime);

                //oForm.Items.Item("tUpDt").Enabled = false;
                //oForm.Items.Item("tUpDt").Click();
                //oForm.Items.Item("tUpTm").Enabled = false;

                Utils.releaseObject(dtSource);
            }
            catch (Exception ex)
            {
                oSBOApplication.MessageBox(ex.Message);
            }
            finally
            {
                oForm.Freeze(false);
            }
        }

        /// <summary>
        /// Item Event update bom version
        /// </summary>
        public void ItemEvent_UpdateBomVer(string formUID, ref ItemEvent pVal, ref bool bubbleEvent)
        {
            switch (pVal.EventType)
            {
                case BoEventTypes.et_CLICK: ItemEvent_UpdateBomVer_Click(formUID, ref pVal, ref bubbleEvent); break;
            }
        }

        private void ItemEvent_UpdateBomVer_Click(string formUID, ref ItemEvent pVal, ref bool bubbleEvent)
        {
            switch (pVal.ItemUID)
            {
                case "btnUpdate": UpdateBomVer_Update(formUID, ref pVal, ref bubbleEvent); break;
            }
        }

        private void UpdateBomVer_Update(string formUID, ref ItemEvent pVal, ref bool bubbleEvent)
        {
            if (bubbleEvent)
            {
                if (pVal.BeforeAction == false && pVal.ActionSuccess == true)
                {
                    SAPbobsCOM.GeneralService oGeneralService;
                    SAPbobsCOM.GeneralData oGeneralData;
                    SAPbobsCOM.GeneralDataCollection oSons;
                    SAPbobsCOM.GeneralData oSon;
                    SAPbobsCOM.CompanyService sCmp;
                    sCmp = oSBOCompany.GetCompanyService();

                    // Get a handle to the UDO
                    oGeneralService = sCmp.GetGeneralService("BOMVER");

                    Recordset oRecBomSap = oSBOCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
                    Recordset oRecBomVer = oSBOCompany.GetBusinessObject(BoObjectTypes.BoRecordset);

                    try
                    {
                        string query = string.Empty;
                        if (oSBOCompany.DbServerType == BoDataServerTypes.dst_HANADB)
                        {
                            query = "CALL SOL_SP_UPDTBOM_GETDIFFBOM()";
                        }

                        oRecBomSap.DoQuery(query);
                        if (oRecBomSap.RecordCount > 0)
                        {
                            for (int i = 1; i <= oRecBomSap.RecordCount; i++)
                            {
                                string test = oRecBomSap.Fields.Item("CodeFgSap").Value;
                                // Get BOM SAP and insert to BOM Version
                                if (oSBOCompany.DbServerType == BoDataServerTypes.dst_HANADB)
                                {
                                    query = "CALL SOL_SP_UPDTBOM_GETBOM('" + oRecBomSap.Fields.Item("CodeFgSap").Value + "')";
                                }
                                oRecBomVer.DoQuery(query);
                                if (oRecBomVer.RecordCount > 0)
                                {
                                    // Specify data for main UDO
                                    oGeneralData = oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralData);
                                    oGeneralData.SetProperty("U_SOL_ITEMCODE", oRecBomVer.Fields.Item("ItemCode").Value);
                                    oGeneralData.SetProperty("U_SOL_ITEMNAME", oRecBomVer.Fields.Item("ItemName").Value);
                                    oGeneralData.SetProperty("U_SOL_ITMGRPCOD", oRecBomVer.Fields.Item("ItmsGrpCod").Value);
                                    oGeneralData.SetProperty("U_SOL_ITEMGROUP", oRecBomVer.Fields.Item("ItmsGrpNam").Value);
                                    oGeneralData.SetProperty("U_SOL_BOMTYPE", oRecBomVer.Fields.Item("BomType").Value);
                                    oGeneralData.SetProperty("U_SOL_VERSION", GenerateBomVersion(oRecBomVer.Fields.Item("ItemCode").Value));
                                    oGeneralData.SetProperty("U_SOL_PLANQTY", oRecBomVer.Fields.Item("PlanQty").Value);
                                    oGeneralData.SetProperty("U_SOL_POSTDATE", DateTime.Now.ToString("yyyyMMdd", CultureInfo.InvariantCulture));
                                    oGeneralData.SetProperty("U_SOL_STATUS", "A");
                                    oGeneralData.SetProperty("U_SOL_WHSCODE", oRecBomVer.Fields.Item("ToWH").Value);
                                    oGeneralData.SetProperty("U_SOL_REMARK", "Created by Add-On Update Bom Version");


                                    oSons = oGeneralData.Child("SOL_BOMVER_D");
                                    while (!oRecBomVer.EoF)
                                    {
                                        oSon = oSons.Add();
                                        oSon.SetProperty("U_SOL_TYPE", oRecBomVer.Fields.Item("Type").Value);
                                        oSon.SetProperty("U_SOL_ITEMCODE", oRecBomVer.Fields.Item("ItemCodeComp").Value);
                                        oSon.SetProperty("U_SOL_ITEMNAME", GetItemName(oRecBomVer.Fields.Item("ItemCodeComp").Value));
                                        oSon.SetProperty("U_SOL_QTY", oRecBomVer.Fields.Item("Quantity").Value);
                                        oSon.SetProperty("U_SOL_WHSCODE", oRecBomVer.Fields.Item("Warehouse").Value);
                                        oSon.SetProperty("U_SOL_UOM", oRecBomVer.Fields.Item("Uom").Value);
                                        oSon.SetProperty("U_SOL_METHOD", oRecBomVer.Fields.Item("IssueMthd").Value);

                                        oRecBomVer.MoveNext();
                                    }
                                    oGeneralService.Add(oGeneralData);
                                }
                                oRecBomSap.MoveNext();
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        oSBOApplication.MessageBox(ex.Message);
                    }
                    finally
                    {
                        Utils.releaseObject(oRecBomSap);
                        Utils.releaseObject(oRecBomVer);
                        Utils.releaseObject(oGeneralService);
                    }
                }
            }
        }

        /// <summary>
        /// Get last update data
        /// </summary>
        private void GetLastUpdate(out string updateDate, out string updateTime)
        {
            updateDate = DateTime.Now.ToString("yyyyMMdd", CultureInfo.InvariantCulture);
            updateTime = DateTime.Now.ToString("HH:mm:ss", CultureInfo.InvariantCulture);

            Recordset oRec = oSBOCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
            string query = "SELECT MAX(\"Code\"), \"U_SOL_UPDATE\", \"U_SOL_UPTIME\" "
                            + "FROM \"@SOL_UPBOMVER_H\" "
                            + "GROUP BY \"Code\", \"U_SOL_UPDATE\", \"U_SOL_UPTIME\" ";
            oRec.DoQuery(query);
            if (oRec.RecordCount > 0)
            {
                updateDate = oRec.Fields.Item("U_SOL_UPDATE").Value;
                updateTime = oRec.Fields.Item("U_SOL_UPTIME").Value;
            }

            Utils.releaseObject(oRec);
        }

        /// <summary>
        /// Generate code number
        /// </summary>
        private int GenerateCode()
        {
            int code = 0;

            Recordset oRec = oSBOCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
            string query = "CALL SOL_SP_UPDTBOM_CODE";
            oRec.DoQuery(query);

            if (oRec.RecordCount > 0)
            {
                code = oRec.Fields.Item("RunNumber").Value;
            }

            Utils.releaseObject(oRec);
            return code;
        }

        /// <summary>
        /// Generate bom number version
        /// </summary>
        private string GenerateBomVersion(string itemCode)
        {
            string version = string.Empty;

            Recordset oRec = oSBOCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
            string query = "CALL SOL_SP_BOMVER_VERSION_CODE('" + itemCode + "')";

            oRec.DoQuery(query);

            if (oRec.RecordCount > 0)
            {
                version = oRec.Fields.Item("RunNumber").Value;
            }

            Utils.releaseObject(oRec);
            return version;
        }

        private string GetItemName(string itemCode)
        {
            string itemName = string.Empty;
            Recordset oRec = oSBOCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
            string query = "SELECT \"ItemName\" FROM OITM WHERE \"ItemCode\" = '" + itemCode + "'";
            oRec.DoQuery(query);

            if (oRec.RecordCount > 0)
            {
                itemName = oRec.Fields.Item("ItemName").Value;
            }

            Utils.releaseObject(oRec);
            return itemName;
        }
    }
}
