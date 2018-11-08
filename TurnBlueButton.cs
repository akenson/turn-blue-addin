using System;
using System.Windows.Forms;
using System.Drawing;
using System.Runtime.InteropServices;
using Inventor;

namespace TurnBlueAddIn
{
    /// <summary>
    /// TurnBlueButton class
    /// </summary>
    internal class TurnBlueUI
    {
        private TurnBlueButton m_TurnBlueButton;

        //user interface event
        private UserInterfaceEvents m_userInterfaceEvents;

        Inventor.Application m_inventorApplication;

        public TurnBlueUI()
        {

        }

        public void InitUI(AddInServer addIn, bool firstTime, Inventor.Application inventorApplication)
        {
            m_inventorApplication = inventorApplication;
            Button.InventorApplication = m_inventorApplication;

            //load image icons for UI items
            Icon turnBlueIcon = new Icon(this.GetType(), "TurnBlue.ico");

            //retrieve the GUID for this class
            GuidAttribute addInCLSID;
            addInCLSID = (GuidAttribute)GuidAttribute.GetCustomAttribute(typeof(AddInServer), typeof(GuidAttribute));
            string addInCLSIDString;
            addInCLSIDString = "{" + addInCLSID.Value + "}";

            m_TurnBlueButton = new TurnBlueButton(addIn,
                "TurnBlue", "Autodesk:TurnBlueAddIn:TurnBlueCmdBtn", CommandTypesEnum.kShapeEditCmdType,
                addInCLSIDString, "Turns the part blue",
                "TurnBlue", turnBlueIcon, turnBlueIcon, ButtonDisplayEnum.kDisplayTextInLearningMode);

            if (firstTime == true)
            {
                //access user interface manager
                UserInterfaceManager userInterfaceManager;
                userInterfaceManager = m_inventorApplication.UserInterfaceManager;

                InterfaceStyleEnum interfaceStyle;
                interfaceStyle = userInterfaceManager.InterfaceStyle;


                //get the ribbon associated with part document
                Inventor.Ribbons ribbons;
                ribbons = userInterfaceManager.Ribbons;

                Inventor.Ribbon partRibbon;
                partRibbon = ribbons["Part"];

                //get the tabs associated with part ribbon
                RibbonTabs ribbonTabs;
                ribbonTabs = partRibbon.RibbonTabs;

                RibbonTab partViewRibbonTab;
                partViewRibbonTab = ribbonTabs["id_TabView"];

                //create a new panel with the tab
                RibbonPanels ribbonPanels;
                ribbonPanels = partViewRibbonTab.RibbonPanels;

                RibbonPanel appearancePanel = ribbonPanels["id_PanelA_ViewAppearance"];

                CommandControls panelCtrls = appearancePanel.CommandControls;
                CommandControl TurnBlueCmdBtnCmdCtrl;
                TurnBlueCmdBtnCmdCtrl = panelCtrls.AddButton(m_TurnBlueButton.ButtonDefinition, false, true, "", false);

            }
        }
    }

    internal class TurnBlueButton : Button
    {
        AddInServer m_addInServer;
      

        #region "Methods"

        public TurnBlueButton(AddInServer addInServer, string displayName, string internalName, CommandTypesEnum commandType, string clientId, string description, string tooltip, Icon standardIcon, Icon largeIcon, ButtonDisplayEnum buttonDisplayType)
            : base(displayName, internalName, commandType, clientId, description, tooltip, standardIcon, largeIcon, buttonDisplayType)
        {
            m_addInServer = addInServer;
        }

        public TurnBlueButton(AddInServer addInServer, string displayName, string internalName, CommandTypesEnum commandType, string clientId, string description, string tooltip, ButtonDisplayEnum buttonDisplayType)
            : base(displayName, internalName, commandType, clientId, description, tooltip, buttonDisplayType)
        {
            m_addInServer = addInServer;
        }

        override protected void ButtonDefinition_OnExecute(NameValueMap context)
        {
            try
            {
                m_addInServer.TurnBlue();
            }
            catch (Exception e)
            {
                MessageBox.Show(e.ToString());
            }
        }
        #endregion

      
    }
}
