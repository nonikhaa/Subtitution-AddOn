using SAPbouiCOM;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Subtitution.Processor;

namespace Subtitution
{
    public class SBOEventHandler
    {
        private SAPbouiCOM.Application oSBOApplication;
        private SAPbobsCOM.Company oSBOCompany;
        public SAPbobsCOM.CompanyService oCompService;

        /// <summary>
        /// Constructor --> first initialization when class is called
        /// </summary>
        #region Constructor
        public SBOEventHandler()
        {

        }

        public SBOEventHandler(SAPbouiCOM.Application oSBOApplication)
        {
            this.oSBOApplication = oSBOApplication;
        }

        public SBOEventHandler(SAPbouiCOM.Application oSBOApplication, SAPbobsCOM.Company oSBOCompany)
        {
            this.oSBOApplication = oSBOApplication;
            this.oSBOCompany = oSBOCompany;
        }
        #endregion

        /// <summary>
        /// Handle app event
        /// </summary>
        public void HandleAppEvent(SAPbouiCOM.BoAppEventTypes EventType)
        {
            switch (EventType)
            {
                case SAPbouiCOM.BoAppEventTypes.aet_ShutDown:
                    if (oSBOCompany.Connected) oSBOCompany.Disconnect();
                    System.Windows.Forms.Application.Exit();
                    break;
                case SAPbouiCOM.BoAppEventTypes.aet_CompanyChanged:
                    if (oSBOCompany.Connected) oSBOCompany.Disconnect();
                    System.Windows.Forms.Application.Exit();
                    break;
                case SAPbouiCOM.BoAppEventTypes.aet_LanguageChanged:
                    if (oSBOCompany.Connected) oSBOCompany.Disconnect();
                    System.Windows.Forms.Application.Exit();
                    break;
                case SAPbouiCOM.BoAppEventTypes.aet_ServerTerminition:
                    if (oSBOCompany.Connected) oSBOCompany.Disconnect();
                    System.Windows.Forms.Application.Exit();
                    break;
            }
        }

        /// <summary>
        /// Handle menu event
        /// </summary>
        public void HandleMenuEvent(ref SAPbouiCOM.MenuEvent pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            AlternativeItem altItem = new AlternativeItem(oSBOApplication, oSBOCompany);
            ChangeCompItem chgCompItem = new ChangeCompItem(oSBOApplication, oSBOCompany);
            UpdateBomVer updtBomVer = new UpdateBomVer(oSBOApplication, oSBOCompany);

            try
            {
                switch (pVal.MenuUID)
                {
                    case "89111": altItem.MenuEvent_AltItem(ref pVal, out BubbleEvent); break;
                    case "89222": chgCompItem.MenuEvent_ChangeCompItm(ref pVal, out BubbleEvent); break;
                    case "89333": updtBomVer.MenuEvent_ChangeCompItm(ref pVal, out BubbleEvent); break;

                    case "AltItemAdd": altItem.MenuEvent_AltItemAdd(ref pVal, ref BubbleEvent); break;
                    case "AltItemDel": altItem.MenuEvent_AltItemDel(ref pVal, ref BubbleEvent); break;
                }
            }
            catch (Exception ex)
            {
                BubbleEvent = false;
                oSBOApplication.StatusBar.SetText(ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }

        /// <summary>
        /// Handle right click event
        /// </summary>
        public void HandleRightClickEvent(ref ContextMenuInfo eventInfo, out bool bubbleEvent)
        {
            Form oForm = oSBOApplication.Forms.ActiveForm;

            try
            {
                bubbleEvent = true;
                AlternativeItem altItem = new AlternativeItem(oSBOApplication, oSBOCompany);

                switch (oForm.TypeEx)
                {
                    case "ALTITEM": altItem.RightClickEvent_AltItem(ref eventInfo, ref bubbleEvent); break;
                }
            }
            catch(Exception ex)
            {
                bubbleEvent = false;
                Utils.releaseObject(oForm);
                oSBOApplication.StatusBar.SetText(ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            }
        }

        /// <summary>
        /// Handle item event
        /// </summary>
        public void HandleItemEvent(string FormUID, ref ItemEvent pVal, out bool bubbleEvent)
        {
            bubbleEvent = true;
            AlternativeItem altItem = new AlternativeItem(oSBOApplication, oSBOCompany);
            ChangeCompItem changeCompItm = new ChangeCompItem(oSBOApplication, oSBOCompany);

            try
            {
                if (pVal.EventType != BoEventTypes.et_FORM_UNLOAD)
                {
                    switch (pVal.FormTypeEx)
                    {
                        case "ALTITEM": altItem.ItemEvent_AltItem(FormUID, ref pVal, ref bubbleEvent); break;
                        case "COMPITM": changeCompItm.ItemEvent_ChangeCompItm(FormUID, ref pVal, ref bubbleEvent); break;
                    }
                }
            }
            catch (Exception ex)
            {
                bubbleEvent = false;
                oSBOApplication.MessageBox(ex.Message);
            }
        }


        /// <summary>
        /// Event when click add menu (CTRL+A)
        /// </summary>
        private void MenuEventHandlerAdd(ref SAPbobsCOM.Company oSBOCompany, ref Application oSBOApplicaton
                                        , ref SAPbouiCOM.MenuEvent pVal, out bool bubbleEvent)
        {
            bubbleEvent = true;
        }
    }
}
