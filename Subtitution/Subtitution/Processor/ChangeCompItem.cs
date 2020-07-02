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

        #region Item Event
        /// <summary>
        /// Item Event change component item
        /// </summary>
        public void ItemEvent_ChangeCompItm(string formUID, ref ItemEvent pVal, ref bool bubbleEvent)
        {
            switch (pVal.EventType)
            {
                case BoEventTypes.et_VALIDATE: Validate_CompItem(formUID, ref pVal, ref bubbleEvent); break;
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
        #endregion
    }
}
