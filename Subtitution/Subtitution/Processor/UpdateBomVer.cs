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
        private ProgressBar oProgressBar;

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

                DateTime updDate = DateTime.Now.Date;
                string updTime = DateTime.Now.TimeOfDay.ToString();
                GetLastUpdate(out updDate, out updTime);

                dtSource.SetValue("U_SOL_UPDATE", 0, updDate.ToString("yyyyMMdd", CultureInfo.InvariantCulture));
                dtSource.SetValue("U_SOL_UPTIME", 0, updTime);

                EditText oEdit = oForm.Items.Item("tEdit").Specific;
                oEdit.Active = true;

                oForm.Items.Item("tUpDt").Enabled = false;
                oForm.Items.Item("tUpTm").Enabled = false;

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

        /// <summary>
        /// Update bom masal
        /// </summary>
        private void UpdateBomVer_Update(string formUID, ref ItemEvent pVal, ref bool bubbleEvent)
        {
            if (bubbleEvent)
            {
                if (pVal.BeforeAction == false && pVal.ActionSuccess == true)
                {
                    SAPbobsCOM.GeneralService oGeneralService;
                    SAPbobsCOM.GeneralData oGeneralData;
                    SAPbobsCOM.CompanyService sCmp;
                    SAPbobsCOM.ProductTrees oBom = (SAPbobsCOM.ProductTrees)oSBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oProductTrees);
                    sCmp = oSBOCompany.GetCompanyService();

                    SAPbobsCOM.GeneralData oGenDataUpdtBom;
                    SAPbobsCOM.GeneralService oGenServiceUpdtBom;

                    Form oForm = oSBOApplication.Forms.Item(formUID);
                    DBDataSource dtSource = null;
                    dtSource = oForm.DataSources.DBDataSources.Item("@SOL_UPBOMVER_H");
                    string lastDate = dtSource.GetValue("U_SOL_UPDATE", 0);
                    string lastTime = Convert.ToString(int.Parse(dtSource.GetValue("U_SOL_UPTIME", 0).Replace(":", "")));

                    // Get a handle to the UDO
                    oGeneralService = sCmp.GetGeneralService("BOMVER");

                    Recordset oRecBomSap = oSBOCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
                    Recordset oRecBomVer = oSBOCompany.GetBusinessObject(BoObjectTypes.BoRecordset);

                    oForm.Freeze(true);

                    try
                    {
                        if (!oSBOCompany.InTransaction)
                        {
                            oSBOCompany.StartTransaction();
                        }

                        string query = string.Empty;
                        if (oSBOCompany.DbServerType == BoDataServerTypes.dst_HANADB)
                        {
                            query = "CALL SOL_SP_UPDTBOM_GETDIFFBOM('" + lastDate + "', '" + lastTime + "')";
                        }
                        oRecBomSap.DoQuery(query);

                        if (oRecBomSap.RecordCount > 0)
                        {
                            int progress = 0;
                            oProgressBar = oSBOApplication.StatusBar.CreateProgressBar("Update BOM", oRecBomSap.RecordCount, true);
                            oProgressBar.Text = "Update BOM...";

                            for (int i = 1; i <= oRecBomSap.RecordCount; i++)
                            {
                                // Get BOM SAP and insert to BOM Version
                                string itemCodeFG = oRecBomSap.Fields.Item("Code").Value;
                                string version = oRecBomSap.Fields.Item("Version").Value;

                                if (oSBOCompany.DbServerType == BoDataServerTypes.dst_HANADB)
                                {
                                    query = "CALL SOL_SP_UPDTBOM_GETBOM('" + itemCodeFG + "')";
                                }
                                oRecBomVer.DoQuery(query);

                                if (oRecBomVer.RecordCount > 0)
                                {
                                    oGeneralData = oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralData);
                                    if (!string.IsNullOrEmpty(version)) // jika bom sudah ada di bom version
                                    {
                                        // non aktivin bom version
                                        InactiveBomVer(itemCodeFG);

                                        // update aktivin bom version sesuai versinya
                                        ActivateBomVer(version);

                                        // update versi di bom sap
                                        UpdateBOM(itemCodeFG, version, ref bubbleEvent);

                                    }
                                    else // jika bom belum ada di bom version
                                    {
                                        // non aktivin bom version
                                        InactiveBomVer(itemCodeFG);

                                        // add bom version
                                        string versionCode = string.Empty;
                                        AddBomVer(ref oRecBomVer, ref oGeneralService, ref bubbleEvent, out versionCode);

                                        // update versi di bom sap
                                        UpdateBOM(itemCodeFG, versionCode, ref bubbleEvent);
                                    }
                                }
                                progress += 1;
                                oProgressBar.Value = progress;

                                oRecBomSap.MoveNext();
                            }

                            // Save ke log update bom version
                            #region Save ke log update bom version
                            oGenServiceUpdtBom = sCmp.GetGeneralService("UPBOMVER");
                            oGenDataUpdtBom = oGenServiceUpdtBom.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralData);
                            oGenDataUpdtBom.SetProperty("Code", GenerateCode());
                            oGenDataUpdtBom.SetProperty("U_SOL_UPDATE", DateTime.Now.Date.ToShortDateString());
                            oGenDataUpdtBom.SetProperty("U_SOL_UPTIME", DateTime.Now.ToString("HH:mm:ss", CultureInfo.InvariantCulture));

                            oGenServiceUpdtBom.Add(oGenDataUpdtBom);
                            #endregion

                            oProgressBar.Stop();
                            TemplateLoad(ref oForm);
                            oSBOApplication.StatusBar.SetText("Update BOM Success", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success);
                        }
                        else
                        {
                            oSBOApplication.MessageBox("Tidak ada data yang dapat di update.");
                        }

                        oSBOCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit);
                    }
                    catch (Exception ex)
                    {
                        bubbleEvent = false;
                        oSBOApplication.MessageBox(ex.Message);
                        oSBOCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
                    }
                    finally
                    {
                        if (oProgressBar != null)
                        {
                            oProgressBar.Stop();
                            Utils.releaseObject(oProgressBar);
                        }

                        if (oForm != null) oForm.Freeze(false);

                        Utils.releaseObject(oRecBomSap);
                        Utils.releaseObject(oRecBomVer);
                        Utils.releaseObject(oGeneralService);
                    }
                }
            }
        }

        /// <summary>
        /// Inactive bom version
        /// </summary>
        private void InactiveBomVer(string itemCodeFG)
        {
            SAPbobsCOM.GeneralService oGeneralService;
            SAPbobsCOM.GeneralData oGeneralData;
            SAPbobsCOM.GeneralDataParams oGeneralParams;
            SAPbobsCOM.CompanyService sCmp;
            Recordset oRecBomVer1 = oSBOCompany.GetBusinessObject(BoObjectTypes.BoRecordset);

            sCmp = oSBOCompany.GetCompanyService();
            oGeneralService = sCmp.GetGeneralService("BOMVER");

            oRecBomVer1.DoQuery("CALL SOL_SP_UPDTBOM_GETACTIVEBOM('" + itemCodeFG + "')");
            if (oRecBomVer1.RecordCount > 0)
            {
                while (!oRecBomVer1.EoF)
                {
                    oGeneralParams = (GeneralDataParams)oGeneralService.GetDataInterface(GeneralServiceDataInterfaces.gsGeneralDataParams);
                    var a = oRecBomVer1.Fields.Item("DocEntry").Value;
                    oGeneralParams.SetProperty("DocEntry", a);

                    oGeneralData = oGeneralService.GetByParams(oGeneralParams);

                    oGeneralData.SetProperty("U_SOL_STATUS", "I");
                    oGeneralService.Update(oGeneralData);

                    oRecBomVer1.MoveNext();
                }
            }
        }

        /// <summary>
        /// Activate bom version
        /// </summary>
        private void ActivateBomVer(string version)
        {
            SAPbobsCOM.GeneralService oGeneralService;
            SAPbobsCOM.GeneralData oGeneralData;
            SAPbobsCOM.GeneralDataParams oGeneralParams;
            SAPbobsCOM.CompanyService sCmp;
            Recordset oRec = oSBOCompany.GetBusinessObject(BoObjectTypes.BoRecordset);

            sCmp = oSBOCompany.GetCompanyService();
            oGeneralService = sCmp.GetGeneralService("BOMVER");

            oRec.DoQuery("SELECT \"DocEntry\" FROM \"@SOL_BOMVER_H\" WHERE \"U_SOL_VERSION\" = '" + version + "'");
            if (oRec.RecordCount > 0)
            {
                while (!oRec.EoF)
                {
                    oGeneralParams = oGeneralService.GetDataInterface(GeneralServiceDataInterfaces.gsGeneralDataParams);
                    var a = oRec.Fields.Item("DocEntry").Value;
                    oGeneralParams.SetProperty("DocEntry", a);

                    oGeneralData = oGeneralService.GetByParams(oGeneralParams);

                    oGeneralData.SetProperty("U_SOL_STATUS", "A");
                    oGeneralService.Update(oGeneralData);

                    oRec.MoveNext();
                }
            }
        }

        /// <summary>
        /// Add new bom version
        /// </summary>
        private void AddBomVer(ref Recordset oRecBomVer, ref GeneralService oGeneralService, ref bool bubbleEvent, out string versionCode)
        {
            SAPbobsCOM.GeneralData oGeneralData;
            SAPbobsCOM.CompanyService sCmp;
            SAPbobsCOM.GeneralDataCollection oSons;
            SAPbobsCOM.GeneralData oSon;

            sCmp = oSBOCompany.GetCompanyService();
            oGeneralService = sCmp.GetGeneralService("BOMVER");
            oGeneralData = oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralData);
            oSons = oGeneralData.Child("SOL_BOMVER_D");

            versionCode = string.Empty;

            try
            {
                // Specify data for main UDO
                versionCode = GenerateBomVersion(oRecBomVer.Fields.Item("ItemCode").Value);
                oGeneralData.SetProperty("U_SOL_ITEMCODE", oRecBomVer.Fields.Item("ItemCode").Value);
                oGeneralData.SetProperty("U_SOL_ITEMNAME", oRecBomVer.Fields.Item("ItemName").Value);
                oGeneralData.SetProperty("U_SOL_ITMGRPCOD", oRecBomVer.Fields.Item("ItmsGrpCod").Value);
                oGeneralData.SetProperty("U_SOL_ITEMGROUP", oRecBomVer.Fields.Item("ItmsGrpNam").Value);
                oGeneralData.SetProperty("U_SOL_BOMTYPE", oRecBomVer.Fields.Item("BomType").Value);
                oGeneralData.SetProperty("U_SOL_VERSION", versionCode);
                oGeneralData.SetProperty("U_SOL_PLANQTY", oRecBomVer.Fields.Item("PlanQty").Value);
                oGeneralData.SetProperty("U_SOL_POSTDATE", DateTime.Now.Date);
                oGeneralData.SetProperty("U_SOL_STATUS", "A");
                oGeneralData.SetProperty("U_SOL_WHSCODE", oRecBomVer.Fields.Item("ToWH").Value);
                oGeneralData.SetProperty("U_SOL_REMARK", "Created by Add-On Update Bom Version");

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
            catch (Exception ex)
            {
                bubbleEvent = false;
                oSBOApplication.MessageBox(ex.Message);
            }
            finally
            {
                Utils.releaseObject(oGeneralData);
                Utils.releaseObject(sCmp);
                Utils.releaseObject(oSons);
            }
        }

        /// <summary>
        /// Update production order
        /// </summary>
        private void UpdateBOM(string itemCodeFg, string version, ref bool bubbleEvent)
        {
            SAPbobsCOM.ProductTrees oBom = (SAPbobsCOM.ProductTrees)oSBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oProductTrees);

            try
            {
                oBom.GetByKey(itemCodeFg);
                oBom.UserFields.Fields.Item("U_SOL_BOMVERNO").Value = version;

                if (oBom.Update() != 0)
                {
                    int ErrCod = oSBOCompany.GetLastErrorCode();
                    string ErrMsg = oSBOCompany.GetLastErrorDescription();

                    oSBOApplication.MessageBox(ErrCod.ToString() + " : " + ErrMsg);
                }
            }
            catch (Exception ex)
            {
                bubbleEvent = false;
                oSBOApplication.MessageBox(ex.Message);
            }
            finally
            {
                Utils.releaseObject(oBom);
            }
        }

        /// <summary>
        /// Get last update data
        /// </summary>
        private void GetLastUpdate(out DateTime updateDate, out string updateTime)
        {
            updateDate = DateTime.ParseExact("20000110", "yyyyMMdd", CultureInfo.InvariantCulture);
            updateTime = DateTime.Now.ToString("HH:mm:ss", CultureInfo.InvariantCulture);

            Recordset oRec = oSBOCompany.GetBusinessObject(BoObjectTypes.BoRecordset);

            string query = "CALL SOL_SP_UPDTBOM_GETLASTUPDT()";

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
        private string GenerateCode()
        {
            string code = "0";

            Recordset oRec = oSBOCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
            string query = "CALL SOL_SP_UPDTBOM_CODE";
            oRec.DoQuery(query);

            if (oRec.RecordCount > 0)
            {
                code = Convert.ToString(oRec.Fields.Item("RunNumber").Value);
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
