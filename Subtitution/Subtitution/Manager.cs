using SAPbobsCOM;
using SAPbouiCOM;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Subtitution
{
    public class Manager
    {
        public const string addonName = "SBO";

        private SAPbouiCOM.Application oSBOApplication = null;
        private SAPbobsCOM.Company oSBOCompany = null;

        public Manager()
        {
            StartUp();
        }

        private void StartUp()
        {
            try
            {
                SetupApplication();
                NumericSeparators();
                CatchingEvents();
                CreateUDT();
                CreateUDF();
                CreateUDO();
                CreateMenu();
                CreateFolder();
                CreateFMS();
                CreateSP();

                oSBOApplication.StatusBar.SetText(addonName + " Add-On Subtitution Connected", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success);
            }
            catch (Exception ex)
            {
                if (oSBOApplication != null)
                    oSBOApplication.MessageBox(ex.Message);
                else
                    MessageBox.Show(ex.Message);

                System.Windows.Forms.Application.Exit();
            }
        }

        /// <summary>
        /// Connect to SAP
        /// </summary>
        private void SetupApplication()
        {
            SAPbouiCOM.SboGuiApi oSboGuiApi = null;
            string sConnectionString = null;

            oSboGuiApi = new SAPbouiCOM.SboGuiApi();
            sConnectionString = System.Convert.ToString(Environment.GetCommandLineArgs().GetValue(1));
            //sConnectionString = "0030002C0030002C00530041005000420044005F00440061007400650076002C0050004C006F006D0056004900490056";

            oSboGuiApi.Connect(sConnectionString);

            oSBOApplication = oSboGuiApi.GetApplication();
            oSBOCompany = oSBOApplication.Company.GetDICompany();

            if (oSBOCompany.DbServerType == SAPbobsCOM.BoDataServerTypes.dst_HANADB)
                GeneralVariables.SQLHandler = new HANAQueries();
            else
                GeneralVariables.SQLHandler = new SQLQueries();
        }

        /// <summary>
        /// Decimal Separator
        /// </summary>
        private void NumericSeparators()
        {
            SAPbobsCOM.Recordset oRec;

            oRec = oSBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            oRec.DoQuery(GeneralVariables.SQLHandler.SeparatorSQL());

            GeneralVariables.WinDecSep = System.Globalization.CultureInfo.CurrentCulture.NumberFormat.NumberDecimalSeparator;
            GeneralVariables.SBODecSep = oRec.Fields.Item("DecSep").Value.ToString();
            GeneralVariables.SBOThousSep = oRec.Fields.Item("ThousSep").Value.ToString();

            System.Runtime.InteropServices.Marshal.ReleaseComObject(oRec);
            oRec = null;
            GC.Collect();
        }

        /// <summary>
        /// All Event
        /// </summary>
        public void CatchingEvents()
        {
            oSBOApplication.AppEvent += new SAPbouiCOM._IApplicationEvents_AppEventEventHandler(SBOApplication_AppEvent);
            oSBOApplication.MenuEvent += new SAPbouiCOM._IApplicationEvents_MenuEventEventHandler(SBOApplication_MenuEvent);
            oSBOApplication.ItemEvent += new SAPbouiCOM._IApplicationEvents_ItemEventEventHandler(SBOApplication_ItemEvent);
            //oSBOApplication.FormDataEvent += new SAPbouiCOM._IApplicationEvents_FormDataEventEventHandler(SBOApplication_FormDataEvent);
            oSBOApplication.RightClickEvent += new SAPbouiCOM._IApplicationEvents_RightClickEventEventHandler(SBOApplication_RightClickEvent);
        }

        /// <summary>
        /// Create UDT
        /// </summary>
        private void CreateUDT()
        {
            // Alternative item master data
            Utils.CreateUDT(oSBOCompany, "SOL_ALTITEM_H", "Alternative Item - Header", BoUTBTableType.bott_MasterData);
            Utils.CreateUDT(oSBOCompany, "SOL_ALTITEM_D", "Alternative Item - Row", BoUTBTableType.bott_MasterDataLines);

            // Update BOM Version
            Utils.CreateUDT(oSBOCompany, "SOL_UPBOMVER_H", "Update BOM Version", BoUTBTableType.bott_MasterData);

            // Change component log
            Utils.CreateUDT(oSBOCompany, "SOL_COMP_LOG", "Change Component Item Log", BoUTBTableType.bott_MasterData);
        }

        /// <summary>
        /// Create UDF
        /// </summary>
        private void CreateUDF()
        {
            #region Alternative item master data
            // Detail
            Utils.CreateUDF(oSBOCompany, "@SOL_ALTITEM_D", "SOL_ITEMCODE", "Alternative Item Code", BoFieldTypes.db_Alpha, BoFldSubTypes.st_None, 50);
            Utils.CreateUDF(oSBOCompany, "@SOL_ALTITEM_D", "SOL_ITEMNAME", "Alternative Item Name", BoFieldTypes.db_Alpha, BoFldSubTypes.st_None, 100);
            Utils.CreateUDF(oSBOCompany, "@SOL_ALTITEM_D", "SOL_FACTOR", "Factor", BoFieldTypes.db_Numeric, BoFldSubTypes.st_None, 11);
            #endregion

            #region Update BOM Version
            Utils.CreateUDF(oSBOCompany, "@SOL_UPBOMVER_H", "SOL_UPDATE", "Last Update Date", BoFieldTypes.db_Date, BoFldSubTypes.st_None, 10);
            Utils.CreateUDF(oSBOCompany, "@SOL_UPBOMVER_H", "SOL_UPTIME", "Last Update Time", BoFieldTypes.db_Alpha, BoFldSubTypes.st_None, 20);
            #endregion

            #region Change component log
            Utils.CreateUDF(oSBOCompany, "@SOL_COMP_LOG", "SOL_DOCNUM", "No. WO", BoFieldTypes.db_Alpha, BoFldSubTypes.st_None, 15);
            Utils.CreateUDF(oSBOCompany, "@SOL_COMP_LOG", "SOL_RPLDATE", "Replace Date", BoFieldTypes.db_Date, BoFldSubTypes.st_None, 10);
            Utils.CreateUDF(oSBOCompany, "@SOL_COMP_LOG", "SOL_ITEMCODE", "Component Item Code", BoFieldTypes.db_Alpha, BoFldSubTypes.st_None, 50);
            Utils.CreateUDF(oSBOCompany, "@SOL_COMP_LOG", "SOL_ITEMNAME", "Component Item Name", BoFieldTypes.db_Alpha, BoFldSubTypes.st_None, 100);
            Utils.CreateUDF(oSBOCompany, "@SOL_COMP_LOG", "SOL_ITMCODE", "Alternative Item Code", BoFieldTypes.db_Alpha, BoFldSubTypes.st_None, 50);
            Utils.CreateUDF(oSBOCompany, "@SOL_COMP_LOG", "SOL_ITMNAME", "Alternative Item Name", BoFieldTypes.db_Alpha, BoFldSubTypes.st_None, 100);
            Utils.CreateUDF(oSBOCompany, "@SOL_COMP_LOG", "SOL_COMPQTY", "Component Planned Qty", BoFieldTypes.db_Float, BoFldSubTypes.st_Quantity, 11);
            Utils.CreateUDF(oSBOCompany, "@SOL_COMP_LOG", "SOL_ALTQTY", "Alternative Planned Qty", BoFieldTypes.db_Float, BoFldSubTypes.st_Quantity, 11);
            Utils.CreateUDF(oSBOCompany, "@SOL_COMP_LOG", "SOL_SELECTED", "Selected", BoFieldTypes.db_Alpha, BoFldSubTypes.st_None, 1);
            #endregion

            #region Bill of Material
            Utils.CreateUDF(oSBOCompany, "OITT", "SOL_BOMVERNO", "BOM Version No.", BoFieldTypes.db_Alpha, BoFldSubTypes.st_None, 30);
            #endregion
        }

        /// <summary>
        /// Create UDO
        /// </summary>
        #region Create UDO
        private void CreateUDO()
        {
            CreateUDO_AltItem();
            CreateUDO_ChangeCompItem();
            CreateUDO_UpBomVer();
        }

        private void CreateUDO_AltItem()
        {
            Utils.UDOChild child = new Utils.UDOChild("SOL_ALTITEM_D");
            Utils.UDOChild[] childs = { child };

            string[] FormColumnAlias = { "Code", "Name" };
            string[] FormColumnDescription = { "Item Code", "Item Name" };

            SAPbobsCOM.BoYesNoEnum[] FormColumnsEditable = { BoYesNoEnum.tYES, BoYesNoEnum.tYES, BoYesNoEnum.tYES };

            Utils.CreateUDOTemplate(oSBOCompany, "ALTITEM", "Alternative Item Master Data", BoUDOObjType.boud_MasterData, "SOL_ALTITEM_H"
                              , BoYesNoEnum.tYES, BoYesNoEnum.tYES, BoYesNoEnum.tNO, BoYesNoEnum.tNO, BoYesNoEnum.tNO
                              , BoYesNoEnum.tNO, BoYesNoEnum.tNO, BoYesNoEnum.tNO, BoYesNoEnum.tNO, BoYesNoEnum.tYES
                              , "ASOL_ALTITEM_H", null, null, 0, 0, FormColumnAlias, FormColumnDescription, FormColumnAlias
                              , FormColumnDescription, FormColumnsEditable, childs);
        }

        private void CreateUDO_ChangeCompItem()
        {
            Utils.CreateUDOTemplate(oSBOCompany, "COMPITM", "Change Component Item", BoUDOObjType.boud_MasterData, "SOL_COMP_LOG"
                              , BoYesNoEnum.tNO, BoYesNoEnum.tNO, BoYesNoEnum.tNO, BoYesNoEnum.tNO, BoYesNoEnum.tNO
                              , BoYesNoEnum.tNO, BoYesNoEnum.tNO, BoYesNoEnum.tNO, BoYesNoEnum.tNO, BoYesNoEnum.tYES
                              , "ASOL_COMP_LOG", null, null, 0, 0, null, null, null
                              , null, null, null);
        }

        private void CreateUDO_UpBomVer()
        {
            Utils.CreateUDOTemplate(oSBOCompany, "UPBOMVER", "Update BOM Version", BoUDOObjType.boud_MasterData, "SOL_UPBOMVER_H"
                              , BoYesNoEnum.tNO, BoYesNoEnum.tNO, BoYesNoEnum.tNO, BoYesNoEnum.tNO, BoYesNoEnum.tNO
                              , BoYesNoEnum.tNO, BoYesNoEnum.tNO, BoYesNoEnum.tNO, BoYesNoEnum.tNO, BoYesNoEnum.tYES
                              , "ASOL_UPBOMVER_H", null, null, 0, 0, null, null, null
                              , null, null, null);
        }

        #endregion

        /// <summary>
        /// Create Menu
        /// </summary>
        private void CreateMenu()
        {
            Utils.CreateMenuEntryTemplate(ref oSBOApplication, ref oSBOCompany, "89898", "Subtitution Add-On", BoMenuType.mt_POPUP);
            Utils.CreateMenuEntryTemplate(ref oSBOApplication, ref oSBOCompany, "89111", "Alternative Item Master Data", BoMenuType.mt_STRING, 0, "89898");
            Utils.CreateMenuEntryTemplate(ref oSBOApplication, ref oSBOCompany, "89222", "Change Component Item", BoMenuType.mt_STRING, 1, "89898");
            Utils.CreateMenuEntryTemplate(ref oSBOApplication, ref oSBOCompany, "89333", "Update BOM Version", BoMenuType.mt_STRING, 2, "89898");
        }

        /// <summary>
        /// Create folder for put FMS or SP
        /// </summary>
        private void CreateFolder()
        {
            string dir = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location);
            string dirPathSP = dir + @"\SP\";
            string dirPathFMS = dir + @"\FMS\";

            if (!Directory.Exists(dirPathSP))
            {
                Directory.CreateDirectory(dirPathSP);
            }

            if (!Directory.Exists(dirPathFMS))
            {
                Directory.CreateDirectory(dirPathFMS);
            }
        }

        /// <summary>
        /// Formated Search (FMS)
        /// </summary>
        #region Create FMS
        private void CreateFMS()
        {
            Utils.CreateQueryCategory(oSBOCompany, "ADDON - Alternative Item Master Data");
            string queryCategory = "ADDON - Alternative Item Master Data";

            if (oSBOCompany.DbServerType == BoDataServerTypes.dst_HANADB)
            {
                Utils.CreateQueryFromText(oSBOCompany, oSBOApplication, queryCategory, "SOL - ALTERNATIVE ITEM - Item Code", "HANA - SOL - ALTERNATIVE ITEM - Item Code.sql");
                Utils.CreateQueryFromText(oSBOCompany, oSBOApplication, queryCategory, "SOL - CHANGE COMPONENT ITEM - Component Item Code", "HANA - SOL - CHANGE COMPONENT ITEM - Component Item Code.sql");
                Utils.CreateQueryFromText(oSBOCompany, oSBOApplication, queryCategory, "SOL - CHANGE COMPONENT ITEM - Alternative Item Code", "HANA - SOL - CHANGE COMPONENT ITEM - Alternative Item Code.sql");
            }

            Utils.CreateFMS(oSBOCompany, queryCategory, "SOL - ALTERNATIVE ITEM - Item Code", "ALTITEM", "tCode", BoYesNoEnum.tNO, BoYesNoEnum.tNO, BoYesNoEnum.tNO);
            Utils.CreateFMS(oSBOCompany, queryCategory, "SOL - ALTERNATIVE ITEM - Item Code", "ALTITEM", "mt_1", BoYesNoEnum.tNO, BoYesNoEnum.tNO, BoYesNoEnum.tNO, "cItmCd");

            Utils.CreateFMS(oSBOCompany, queryCategory, "SOL - CHANGE COMPONENT ITEM - Component Item Code", "COMPITM", "tComItmCd", BoYesNoEnum.tNO, BoYesNoEnum.tNO, BoYesNoEnum.tNO);
            Utils.CreateFMS(oSBOCompany, queryCategory, "SOL - CHANGE COMPONENT ITEM - Alternative Item Code", "COMPITM", "tAltItmCd", BoYesNoEnum.tNO, BoYesNoEnum.tNO, BoYesNoEnum.tNO);
        }

        #endregion

        #region Create SP
        /// <summary>
        /// Create Stored Procedure (SP) - load from file
        /// </summary>
        private void CreateSP()
        {
            CreateSP_ChangeCompItm();
            CreateSP_UpdateBom();
        }

        /// <summary>
        /// SP - change component item
        /// </summary>
        private void CreateSP_ChangeCompItm()
        {
            try
            {
                if (oSBOCompany.DbServerType == BoDataServerTypes.dst_HANADB)
                {
                    Utils.CreateSP(ref oSBOCompany, ref oSBOApplication, "HANA - SOL_SP_COMPITEM_FIND.sql", "SOL_SP_COMPITEM_FIND");
                    Utils.CreateSP(ref oSBOCompany, ref oSBOApplication, "HANA - SOL_SP_COMPITEM_LOG_CODE.sql", "SOL_SP_COMPITEM_LOG_CODE");
                }
            }
            catch (Exception ex)
            {
                oSBOApplication.MessageBox(ex.Message);
            }
        }

        /// <summary>
        /// SP - Update bom version
        /// </summary>
        private void CreateSP_UpdateBom()
        {
            try
            {
                if (oSBOCompany.DbServerType == BoDataServerTypes.dst_HANADB)
                {
                    Utils.CreateSP(ref oSBOCompany, ref oSBOApplication, "HANA - SOL_SP_UPDTBOM_CODE.sql", "SOL_SP_UPDTBOM_CODE");
                    Utils.CreateSP(ref oSBOCompany, ref oSBOApplication, "HANA - SOL_SP_UPDTBOM_GETDIFFBOM.sql", "SOL_SP_UPDTBOM_GETDIFFBOM");
                    Utils.CreateSP(ref oSBOCompany, ref oSBOApplication, "HANA - SOL_SP_UPDTBOM_GETBOM.sql", "SOL_SP_UPDTBOM_GETBOM");
                    Utils.CreateSP(ref oSBOCompany, ref oSBOApplication, "HANA - SOL_SP_BOMVER_VERSION_CODE.sql", "SOL_SP_BOMVER_VERSION_CODE");
                    Utils.CreateSP(ref oSBOCompany, ref oSBOApplication, "HANA - SOL_SP_UPDTBOM_GETACTIVEBOM.sql", "SOL_SP_UPDTBOM_GETACTIVEBOM");
                    Utils.CreateSP(ref oSBOCompany, ref oSBOApplication, "HANA - SOL_SP_UPDTBOM_GETLASTUPDT.sql", "SOL_SP_UPDTBOM_GETLASTUPDT");
                }
            }
            catch (Exception ex)
            {
                oSBOApplication.MessageBox(ex.Message);
            }
        }

        #endregion

        #region SBO Event Handler
        private void SBOApplication_AppEvent(SAPbouiCOM.BoAppEventTypes EventType)
        {
            SBOEventHandler oSBOEventHandler = new SBOEventHandler(oSBOApplication, oSBOCompany);
            oSBOEventHandler.HandleAppEvent(EventType);
        }

        private void SBOApplication_MenuEvent(ref SAPbouiCOM.MenuEvent pVal, out bool BubbleEvent)
        {
            SBOEventHandler oSBOEventHandler = new SBOEventHandler(oSBOApplication, oSBOCompany);
            oSBOEventHandler.HandleMenuEvent(ref pVal, out BubbleEvent);
        }

        private void SBOApplication_ItemEvent(string FormUID, ref SAPbouiCOM.ItemEvent pVal, out bool BubbleEvent)
        {
            SBOEventHandler oSBOEventHandler = new SBOEventHandler(oSBOApplication, oSBOCompany);
            oSBOEventHandler.HandleItemEvent(FormUID, ref pVal, out BubbleEvent);
        }

        //private void SBOApplication_FormDataEvent(ref SAPbouiCOM.BusinessObjectInfo BusinessObjectInfo, out bool BubbleEvent)
        //{
        //    SBOEventHandler oSBOEventHandler = new SBOEventHandler(oSBOApplication, oSBOCompany);
        //    oSBOEventHandler.HandleFormDataEvent(ref BusinessObjectInfo, out BubbleEvent);
        //}

        private void SBOApplication_RightClickEvent(ref SAPbouiCOM.ContextMenuInfo eventInfo, out bool BubbleEvent)
        {
            SBOEventHandler oSBOEventHandler = new SBOEventHandler(oSBOApplication, oSBOCompany);
            oSBOEventHandler.HandleRightClickEvent(ref eventInfo, out BubbleEvent);
        }
        #endregion
    }
}
