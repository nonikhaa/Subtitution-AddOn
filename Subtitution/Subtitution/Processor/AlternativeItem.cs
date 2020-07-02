using SAPbobsCOM;
using SAPbouiCOM;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Subtitution.Processor
{
    public class AlternativeItem
    {
        private SAPbouiCOM.Application oSBOApplication;
        private SAPbobsCOM.Company oSBOCompany;

        /// <summary>
        /// Constructor
        /// </summary>
        public AlternativeItem(SAPbouiCOM.Application oSBOApplication, SAPbobsCOM.Company oSBOCompany)
        {
            this.oSBOApplication = oSBOApplication;
            this.oSBOCompany = oSBOCompany;
        }

        #region Menu Event
        /// <summary>
        /// Menu Event Alternative Item
        /// When click menu, this event called
        /// </summary>
        public void MenuEvent_AltItem(ref MenuEvent pVal, out bool bubbleEvent)
        {
            bubbleEvent = true;
            if (pVal.BeforeAction == false)
            {
                Form oForm = null;

                try
                {
                    oForm = Utils.createForm(ref oSBOApplication, "Alternative Item Master Data");
                    Template_AltItem(ref oForm);
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

        /// <summary>
        /// Menu add row
        /// </summary>
        public void MenuEvent_AltItemAdd(ref MenuEvent pVal, ref bool bubbleEvent)
        {
            if (pVal.BeforeAction == true)
            {
                Form oForm = oSBOApplication.Forms.ActiveForm;
                if (oForm.TypeEx == "ALTITEM")
                {
                    try
                    {
                        oForm.Freeze(true);

                        string mtxName = "mt_1";
                        string dataSource = "@SOL_ALTITEM_D";

                        Matrix oMtx = oForm.Items.Item(mtxName).Specific;
                        oMtx.FlushToDataSource();

                        DBDataSource dtSource = oForm.DataSources.DBDataSources.Item(dataSource);
                        dtSource.InsertRecord(GeneralVariables.iDelRow);

                        dtSource.SetValue("U_SOL_ITEMCODE", GeneralVariables.iDelRow, "");
                        dtSource.SetValue("U_SOL_ITEMNAME", GeneralVariables.iDelRow, "");
                        dtSource.SetValue("U_SOL_FACTOR", GeneralVariables.iDelRow, "");

                        oMtx.LoadFromDataSource();
                        for (int i = 1; i <= oMtx.RowCount; i++)
                        {
                            oMtx.Columns.Item("#").Cells.Item(i).Specific.value = i;
                        }
                        oMtx.FlushToDataSource();

                        oMtx.Columns.Item(1).Cells.Item(GeneralVariables.iDelRow + 1).Click();
                        Utils.releaseObject(dtSource);
                        Utils.releaseObject(oMtx);
                    }
                    catch (Exception ex)
                    {
                        oSBOApplication.MessageBox(ex.Message);
                    }
                    finally
                    {
                        oForm.Freeze(false);
                        Utils.releaseObject(oForm);
                    }
                }
            }
        }

        /// <summary>
        /// Menu delete row
        /// </summary>
        public void MenuEvent_AltItemDel(ref MenuEvent pVal, ref bool bubbleEvent)
        {
            if (pVal.BeforeAction == true)
            {
                Form oForm = oSBOApplication.Forms.ActiveForm;
                if (oForm.TypeEx == "ALTITEM")
                {
                    try
                    {
                        oForm.Freeze(true);

                        string mtxName = "mt_1";
                        string dataSource = "@SOL_ALTITEM_D";

                        Matrix oMtx = oForm.Items.Item(mtxName).Specific;
                        oMtx.FlushToDataSource();

                        try
                        {
                            oMtx.DeleteRow(GeneralVariables.iDelRow);
                        }
                        catch (Exception ex) { }

                        for (int i = 1; i <= oMtx.RowCount; i++)
                        {
                            oMtx.Columns.Item("#").Cells.Item(i).Specific.value = i;
                        }
                        RefreshMatrix(oForm, mtxName, dataSource);

                        Utils.releaseObject(oMtx);

                    }
                    catch (Exception ex)
                    {
                        oSBOApplication.MessageBox(ex.Message);
                    }
                    finally
                    {
                        oForm.Freeze(false);
                        Utils.releaseObject(oForm);
                    }
                }
            }
        }

        /// <summary>
        /// Refresh Matrix
        /// </summary>
        private void RefreshMatrix(Form oForm, string mtxName, string dataSource)
        {
            DBDataSource dtSource = oForm.DataSources.DBDataSources.Item(dataSource);
            Matrix oMtx = oForm.Items.Item(mtxName).Specific;
            dtSource.Clear();
            oMtx.FlushToDataSource();
            Utils.releaseObject(dtSource);
            Utils.releaseObject(oMtx);
        }

        private void Template_AltItem(ref Form oForm)
        {
            DBDataSource altItem_D = oForm.DataSources.DBDataSources.Item("@SOL_ALTITEM_D");
            altItem_D.Clear();
            altItem_D.InsertRecord(altItem_D.Size);

            altItem_D.Offset = altItem_D.Size - 1;
            altItem_D.SetValue("LineId", altItem_D.Size - 1, altItem_D.Size.ToString());

            oForm.Items.Item("mt_1").Specific.LoadFromDataSource();

            Utils.releaseObject(altItem_D);
        }

        #endregion

        #region Item Event
        /// <summary>
        /// Item event
        /// </summary>
        public void ItemEvent_AltItem(string formUID, ref ItemEvent pVal, ref bool bubbleEvent)
        {
            switch (pVal.EventType)
            {
                case BoEventTypes.et_VALIDATE: Validate_AltItem(formUID, ref pVal, ref bubbleEvent); break;
            }
        }

        /// <summary>
        /// Validate alternative item
        /// </summary>
        private void Validate_AltItem(string formUID, ref ItemEvent pVal, ref bool bubbleEvent)
        {
            if (pVal.ItemUID == "tCode")
                Validate_AltItem_CompItemCode(formUID, ref pVal, ref bubbleEvent);
            else if (pVal.ColUID == "cItmCd")
                Validate_AltItem_AltItemCode(formUID, ref pVal, ref bubbleEvent);
        }

        /// <summary>
        /// Header
        /// Show item name when item code choosed
        /// </summary>
        private void Validate_AltItem_CompItemCode(string formUID, ref ItemEvent pVal, ref bool bubbleEvent)
        {
            if (bubbleEvent)
            {
                if (pVal.BeforeAction == false && pVal.ActionSuccess == true && pVal.ItemChanged == true)
                {
                    Form oForm = oSBOApplication.Forms.Item(formUID);
                    DBDataSource dtSource = oForm.DataSources.DBDataSources.Item("@SOL_ALTITEM_H");
                    Recordset oRec = oSBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                    String itemCode = oForm.Items.Item("tCode").Specific.Value;
                    try
                    {
                        oForm.Freeze(true);
                        string name = string.Empty;

                        oRec.DoQuery("SELECT \"ItemName\" FROM OITM WHERE \"ItemCode\" = '" + itemCode + "'");
                        if (oRec.RecordCount > 0)
                        {
                            name = oRec.Fields.Item(0).Value;
                        }

                        dtSource.SetValue("Name", 0, name);
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
        /// Row
        /// Show item name when item code choosed
        /// </summary>
        private void Validate_AltItem_AltItemCode(string formUID, ref ItemEvent pVal, ref bool bubbleEvent)
        {
            if (bubbleEvent)
            {
                if (pVal.BeforeAction == false && pVal.ActionSuccess == true && pVal.ItemChanged == true)
                {
                    Form oForm = oSBOApplication.Forms.Item(formUID);
                    Matrix oMtx = oForm.Items.Item("mt_1").Specific;
                    DBDataSource dtSource = oForm.DataSources.DBDataSources.Item("@SOL_ALTITEM_D");
                    Recordset oRec = oSBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                    String itemCode = oMtx.Columns.Item("cItmCd").Cells.Item(pVal.Row).Specific.Value;

                    try
                    {
                        oForm.Freeze(true);
                        oMtx.FlushToDataSource();

                        if (pVal.Row == oMtx.RowCount)
                        {
                            dtSource.InsertRecord(dtSource.Size);
                            dtSource.SetValue("LineId", dtSource.Size - 1, dtSource.Size.ToString());
                        }

                        string name = string.Empty;
                        oRec.DoQuery("SELECT \"ItemName\" FROM OITM WHERE \"ItemCode\" = '" + itemCode + "'");

                        if (oRec.RecordCount > 0)
                        {
                            name = oRec.Fields.Item(0).Value;
                        }

                        dtSource.SetValue("U_SOL_ITEMNAME", pVal.Row - 1, name);
                        oMtx.LoadFromDataSource();
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
                        Utils.releaseObject(oMtx);
                        Utils.releaseObject(dtSource);
                        Utils.releaseObject(oRec);
                    }
                }
            }
        }

        #endregion

        #region Right Click Event
        /// <summary>
        /// Create menu when rihgt click in matrix
        /// menu Add Row and Delete Row
        public void RightClickEvent_AltItem(ref ContextMenuInfo eventInfo, ref bool bubbleEvent)
        {
            Form oForm = oSBOApplication.Forms.ActiveForm;
            if (eventInfo.BeforeAction == true && eventInfo.ItemUID == "mt_1")
            {
                MenuItem oMenuItem = null;
                Menus oMenus = null;
                MenuCreationParams oCreateionPackage = null;

                try
                {
                    oCreateionPackage = oSBOApplication.CreateObject(BoCreatableObjectType.cot_MenuCreationParams);


                    oCreateionPackage.Type = BoMenuType.mt_STRING;
                    oCreateionPackage.UniqueID = "AltItemAdd";
                    oCreateionPackage.Position = 1;
                    oCreateionPackage.String = "Add Row";
                    oCreateionPackage.Enabled = true;

                    oMenuItem = oSBOApplication.Menus.Item("1280");
                    oMenus = oMenuItem.SubMenus;
                    oMenus.AddEx(oCreateionPackage);


                    oCreateionPackage = oSBOApplication.CreateObject(BoCreatableObjectType.cot_MenuCreationParams);


                    oCreateionPackage.Type = BoMenuType.mt_STRING;
                    oCreateionPackage.UniqueID = "AltItemDel";
                    oCreateionPackage.Position = 2;
                    oCreateionPackage.String = "Delete Row";
                    oCreateionPackage.Enabled = true;

                    oMenuItem = oSBOApplication.Menus.Item("1280");
                    oMenus = oMenuItem.SubMenus;
                    oMenus.AddEx(oCreateionPackage);

                }
                catch (Exception ex)
                {
                    oSBOApplication.MessageBox(ex.Message);
                }
                finally
                {
                    Utils.releaseObject(oMenuItem);
                    Utils.releaseObject(oMenus);
                    Utils.releaseObject(oCreateionPackage);
                }

                GeneralVariables.iDelRow = eventInfo.Row;
            }
            else
            {
                oSBOApplication.Menus.RemoveEx("AltItemAdd");
                oSBOApplication.Menus.RemoveEx("AltItemDel");
            }
            Utils.releaseObject(oForm);
        }
        #endregion
    }
}
