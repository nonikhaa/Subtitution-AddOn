using SAPbobsCOM;
using SAPbouiCOM;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Subtitution.Processor
{
    public class ChangeCompItem
    {
        private SAPbouiCOM.Application oSBOApplication;
        private SAPbobsCOM.Company oSBOCompany;

        /// <summary>
        /// Constructor
        /// </summary>
        public ChangeCompItem(SAPbouiCOM.Application oSBOApplication, SAPbobsCOM.Company oSBOCompany)
        {
            this.oSBOApplication = oSBOApplication;
            this.oSBOCompany = oSBOCompany;
        }

        /// <summary>
        /// Menu Event Change Component Item
        /// When click menu, this event called
        /// </summary>
        public void MenuEvent_ChangeCompItm(ref MenuEvent pVal, out bool bubbleEvent)
        {
            bubbleEvent = true;
            if (pVal.BeforeAction == false)
            {
                Form oForm = null;

                try
                {
                    oForm = Utils.createForm(ref oSBOApplication, "Change Component Item");
                    oForm.Visible = true;
                    TemplateLoad(ref oForm);
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
                dtSource = oForm.DataSources.DBDataSources.Item("@SOL_COMP_LOG");
                oForm.Items.Item("mt_1").Specific.LoadFromDataSource();

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

        #region Item Event
        /// <summary>
        /// Item Event change component item
        /// </summary>
        public void ItemEvent_ChangeCompItm(string formUID, ref ItemEvent pVal, ref bool bubbleEvent)
        {
            switch (pVal.EventType)
            {
                case BoEventTypes.et_VALIDATE: Validate_CompItem(formUID, ref pVal, ref bubbleEvent); break;
                case BoEventTypes.et_CLICK: ItemEvent_ChangeCompItm_Click(formUID, ref pVal, ref bubbleEvent); break;
                    //case BoEventTypes.et_DOUBLE_CLICK: ChangeComp_SelectAll(formUID, ref pVal, ref bubbleEvent); break;
            }
        }

        /// <summary>
        /// Validate component item
        /// </summary>
        private void Validate_CompItem(string formUID, ref ItemEvent pVal, ref bool bubbleEvent)
        {
            switch (pVal.ItemUID)
            {
                case "tComItmCd": Validate_CompItem_CompItemCode(formUID, ref pVal, ref bubbleEvent); break;
                case "tAltItmCd": Validate_CompItem_AltItemCode(formUID, ref pVal, ref bubbleEvent); break;
            }
        }

        /// <summary>
        /// Validate component item code
        /// </summary>
        private void Validate_CompItem_CompItemCode(string formUID, ref ItemEvent pVal, ref bool bubbleEvent)
        {
            if (bubbleEvent)
            {
                if (pVal.BeforeAction == false && pVal.ActionSuccess == true && pVal.ItemChanged == true)
                {
                    Form oForm = oSBOApplication.Forms.Item(formUID);
                    DBDataSource dtSource = oForm.DataSources.DBDataSources.Item("@SOL_COMP_LOG");
                    Recordset oRec = oSBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                    String itemCode = oForm.Items.Item("tComItmCd").Specific.Value;
                    try
                    {
                        oForm.Freeze(true);
                        string name = string.Empty;

                        oRec.DoQuery("SELECT \"ItemName\" FROM OITM WHERE \"ItemCode\" = '" + itemCode + "'");
                        if (oRec.RecordCount > 0)
                        {
                            name = oRec.Fields.Item(0).Value;
                        }

                        dtSource.SetValue("U_SOL_ITEMNAME", 0, name);
                    }
                    catch (Exception ex)
                    {
                        bubbleEvent = false;
                        oSBOApplication.MessageBox(ex.Message);
                    }
                    finally
                    {
                        if (oForm != null) oForm.Freeze(false);

                        Utils.releaseObject(oForm);
                        Utils.releaseObject(oRec);
                        Utils.releaseObject(dtSource);
                    }
                }
            }
        }

        /// <summary>
        /// Validate alternative item code
        /// </summary>
        private void Validate_CompItem_AltItemCode(string formUID, ref ItemEvent pVal, ref bool bubbleEvent)
        {
            if (bubbleEvent)
            {
                if (pVal.BeforeAction == false && pVal.ActionSuccess == true && pVal.ItemChanged == true)
                {
                    Form oForm = oSBOApplication.Forms.Item(formUID);
                    DBDataSource dtSource = oForm.DataSources.DBDataSources.Item("@SOL_COMP_LOG");
                    Recordset oRec = oSBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                    String itemCode = oForm.Items.Item("tAltItmCd").Specific.Value;
                    try
                    {
                        oForm.Freeze(true);
                        string name = string.Empty;

                        oRec.DoQuery("SELECT \"ItemName\" FROM OITM WHERE \"ItemCode\" = '" + itemCode + "'");
                        if (oRec.RecordCount > 0)
                        {
                            name = oRec.Fields.Item(0).Value;
                        }

                        dtSource.SetValue("U_SOL_ITMNAME", 0, name);
                    }
                    catch (Exception ex)
                    {
                        bubbleEvent = false;
                        oSBOApplication.MessageBox(ex.Message);
                    }
                    finally
                    {
                        if (oForm != null) oForm.Freeze(false);

                        Utils.releaseObject(oForm);
                        Utils.releaseObject(oRec);
                        Utils.releaseObject(dtSource);
                    }
                }
            }
        }

        private void ItemEvent_ChangeCompItm_Click(string formUID, ref ItemEvent pVal, ref bool bubbleEvent)
        {
            switch (pVal.ItemUID)
            {
                case "btnFind": ChangeCompItm_Find(formUID, ref pVal, ref bubbleEvent); break;
                case "btnReplace": ChangeCompItm_Replace(formUID, ref pVal, ref bubbleEvent); break;
                    //case "chkAll": ChangeComp_SelectAll(formUID, ref pVal, ref bubbleEvent); break;
            }
        }

        /// <summary>
        /// Button find clicked
        /// </summary>
        private void ChangeCompItm_Find(string formUID, ref ItemEvent pVal, ref bool bubbleEvent)
        {
            if (bubbleEvent)
            {
                if (pVal.BeforeAction == false && pVal.ActionSuccess == true)// && pVal.EventType == BoEventTypes.et_CLICK)
                {
                    Form oForm = oSBOApplication.Forms.Item(formUID);
                    oForm.Freeze(true);

                    Recordset oRec = oSBOCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
                    Matrix oMtx = oForm.Items.Item("mt_1").Specific;
                    string compItm = oForm.Items.Item("tComItmCd").Specific.Value;
                    string altItm = oForm.Items.Item("tAltItmCd").Specific.Value;
                    string startDate = oForm.Items.Item("tStrDt").Specific.Value;
                    string endDate = oForm.Items.Item("tEndDt").Specific.Value;

                    ProgressBar oProgressBar = oSBOApplication.StatusBar.CreateProgressBar("Find Work Order", oMtx.RowCount, true);
                    oProgressBar.Text = "Find Work Order...";

                    try
                    {
                        string query = string.Empty;
                        if (oSBOCompany.DbServerType == BoDataServerTypes.dst_HANADB)
                        {
                            query = "CALL SOL_SP_COMPITEM_FIND('" + compItm + "', '" + altItm + "', '" + startDate + "', '" + endDate + "')";
                        }

                        oRec.DoQuery(query);

                        if (oRec.RecordCount > 0)
                        {
                            EnableDisableMatrix(true, ref oMtx);
                            for (int i = 1; i <= oRec.RecordCount; i++)
                            {
                                oMtx.AddRow();
                                int currentRow = oMtx.RowCount;

                                oMtx.Columns.Item("cCheck").Cells.Item(i).Specific.Checked = true;
                                oMtx.Columns.Item("cFgCode").Cells.Item(i).Specific.Value = oRec.Fields.Item("ItemCode").Value;
                                oMtx.Columns.Item("cFgName").Cells.Item(i).Specific.Value = oRec.Fields.Item("ProdName").Value;
                                oMtx.Columns.Item("cNoWo").Cells.Item(i).Specific.Value = oRec.Fields.Item("DocNum").Value;
                                oMtx.Columns.Item("cComQty").Cells.Item(i).Specific.Value = oRec.Fields.Item("PlannedQty").Value;
                                oMtx.Columns.Item("cAltQty").Cells.Item(i).Specific.Value = oRec.Fields.Item("Alternative Qty").Value;
                            }
                            oForm.Items.Item("tAltItmCd").Click();
                            EnableDisableMatrix(false, ref oMtx);
                        }
                    }
                    catch (Exception ex)
                    {
                        oSBOApplication.MessageBox(ex.Message);
                    }
                    finally
                    {
                        oProgressBar.Stop();
                        if (oForm != null) oForm.Freeze(false);

                        Utils.releaseObject(oForm);
                        Utils.releaseObject(oMtx);
                        Utils.releaseObject(oRec);
                    }
                }
            }
        }

        /// <summary>
        /// Enable and disable row matrix
        /// </summary>
        private void EnableDisableMatrix(bool value, ref Matrix oMtx)
        {
            oMtx.Columns.Item("cCheck").Editable = true;
            oMtx.Columns.Item("cNoWo").Editable = value;
            oMtx.Columns.Item("cFgCode").Editable = value;
            oMtx.Columns.Item("cFgName").Editable = value;
            oMtx.Columns.Item("cComQty").Editable = value;
            oMtx.Columns.Item("cAltQty").Editable = value;
        }

        private void ChangeCompItm_Replace(string formUID, ref ItemEvent pVal, ref bool bubbleEvent)
        {
            if (bubbleEvent)
            {
                if (pVal.BeforeAction == false && pVal.ActionSuccess == true)// && pVal.EventType == BoEventTypes.et_CLICK)
                {
                    Form oForm = oSBOApplication.Forms.Item(formUID);
                    oForm.Freeze(true);

                    SAPbobsCOM.ProductionOrders oProd;
                    oProd = oSBOCompany.GetBusinessObject(BoObjectTypes.oProductionOrders);

                    SAPbobsCOM.GeneralService oGenService = oSBOCompany.GetCompanyService().GetGeneralService("COMPITM");
                    SAPbobsCOM.GeneralData compLog = (SAPbobsCOM.GeneralData)oGenService.GetDataInterface(GeneralServiceDataInterfaces.gsGeneralData);

                    Recordset oRec = oSBOCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
                    Matrix oMtx = oForm.Items.Item("mt_1").Specific;

                    ProgressBar oProgressBar = oSBOApplication.StatusBar.CreateProgressBar("Replace Work Order", oMtx.RowCount, true);
                    oProgressBar.Text = "Replace Work Order...";

                    try
                    {
                        for (int i = 1; i <= oMtx.RowCount; i++)
                        {
                            //dtSource.Offset = dtSource.Size - 1;
                            string check = string.Empty;
                            if (oMtx.Columns.Item("cCheck").Cells.Item(i).Specific.Checked == true)
                                check = "Y";

                            if (check == "Y")
                            {
                                string query = "SELECT \"DocEntry\" FROM OWOR WHERE \"DocNum\" = '" + oMtx.Columns.Item("cNoWo").Cells.Item(i).Specific.Value + "'";
                                oRec.DoQuery(query);

                                int docEntry = 0;
                                if (oRec.RecordCount > 0)
                                {
                                    docEntry = oRec.Fields.Item("DocEntry").Value;
                                }

                                oProd.GetByKey(docEntry);

                                for (int j = 0; j < oProd.Lines.Count; j++)
                                {
                                    if (oProd.Lines.ItemNo == oForm.Items.Item("tComItmCd").Specific.Value)
                                    {
                                        oProd.Lines.SetCurrentLine(j);
                                        oProd.Lines.ItemNo = oForm.Items.Item("tAltItmCd").Specific.Value;
                                        oProd.Lines.PlannedQuantity = Utils.SBOToWindowsNumberWithoutCurrency(oMtx.Columns.Item("cAltQty").Cells.Item(i).Specific.Value);

                                        int retCode = oProd.Update();
                                        if (retCode == 0)
                                        {
                                            compLog.SetProperty("Code", GetLogCode());
                                            compLog.SetProperty("U_SOL_DOCNUM", oMtx.Columns.Item("cNoWo").Cells.Item(i).Specific.Value);
                                            compLog.SetProperty("U_SOL_RPLDATE", DateTime.Now.Date);
                                            compLog.SetProperty("U_SOL_ITEMCODE", oForm.Items.Item("tComItmCd").Specific.Value);
                                            compLog.SetProperty("U_SOL_ITEMNAME", oForm.Items.Item("tComItmNm").Specific.Value);
                                            compLog.SetProperty("U_SOL_ITMCODE", oForm.Items.Item("tAltItmCd").Specific.Value);
                                            compLog.SetProperty("U_SOL_ITMNAME", oForm.Items.Item("tAltItmNm").Specific.Value);
                                            compLog.SetProperty("U_SOL_COMPQTY", oMtx.Columns.Item("cComQty").Cells.Item(i).Specific.Value);
                                            compLog.SetProperty("U_SOL_ALTQTY", oMtx.Columns.Item("cAltQty").Cells.Item(i).Specific.Value);

                                            oGenService.Add(compLog);
                                            if (oSBOCompany.InTransaction)
                                            {
                                                oSBOCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit);
                                            }
                                        }
                                        else
                                        {
                                            throw new Exception();
                                        }
                                    }
                                }
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        oSBOApplication.MessageBox(ex.Message);
                    }
                    finally
                    {
                        oProgressBar.Stop();
                        if (oForm != null) oForm.Freeze(false);
                    }
                }
            }
        }

        /// <summary>
        /// Get Component Log Code
        /// </summary>
        /// <returns></returns>
        private string GetLogCode()
        {
            string code = string.Empty;
            string query = string.Empty;

            Recordset oRec = oSBOCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
            query = "CALL SOL_SP_COMPITEM_LOG_CODE";

            oRec.DoQuery(query);
            if(oRec.RecordCount > 0)
            {
                code = oRec.Fields.Item("RunNumber").Value;
            }

            Utils.releaseObject(oRec);
            return code;
        }
        #endregion
    }
}
