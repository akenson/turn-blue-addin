using System;
using System.Runtime.InteropServices;
using System.Drawing;
using System.Windows.Forms;
using Inventor;
using Microsoft.Win32;

namespace TurnBlueAddIn
{
	/// <summary>
	/// This is the primary AddIn Server class that implements the ApplicationAddInServer interface
	/// that all Inventor AddIns are required to implement. The communication between Inventor and
	/// the AddIn is via the methods on this interface.
	/// </summary>

	[GuidAttribute("B99DB61B-F61E-4A56-AE2C-3FB608A2547D")]
	public class AddInServer : Inventor.ApplicationAddInServer
	{
		#region Data Members
		
		//Inventor application object
		private Inventor.Application m_inventorApplication;
        private TurnBlueUI m_turnBlueUI;
		

		#endregion

		public AddInServer()
		{
		}

		#region ApplicationAddInServer Members

		public void Activate(Inventor.ApplicationAddInSite addInSiteObject, bool firstTime)
		{
			try
			{
				//the Activate method is called by Inventor when it loads the addin
				//the AddInSiteObject provides access to the Inventor Application object
				//the FirstTime flag indicates if the addin is loaded for the first time

				//initialize AddIn members
                m_inventorApplication = addInSiteObject.Application;

                // Initialize the UI components
                m_turnBlueUI = new TurnBlueUI();
                m_turnBlueUI.InitUI(this, firstTime, m_inventorApplication);

                MessageBox.Show("Turn Blue Add-in enabled");
			}
			catch(Exception e)
			{
				MessageBox.Show(e.ToString());
			}
		}

        

		public void Deactivate()
		{
			//the Deactivate method is called by Inventor when the AddIn is unloaded
			//the AddIn will be unloaded either manually by the user or
			//when the Inventor session is terminated
		
			try
			{
				//release inventor Application object
				Marshal.ReleaseComObject(m_inventorApplication);
                m_inventorApplication = null;

				GC.WaitForPendingFinalizers();
				GC.Collect();
			}
			catch(Exception e)
			{
				MessageBox.Show(e.ToString());
			}
		}

        public void TurnBlue()
        {
            // Get the active document
            PartDocument doc = (PartDocument) m_inventorApplication.ActiveDocument;
            AssetLibrary lib =  m_inventorApplication.AssetLibraries["Autodesk Appearance Library"];
            Asset libAsset = lib.AppearanceAssets["Blue - Wall Paint - Glossy"];
            Asset localAsset = libAsset.CopyTo(doc);
            doc.ActiveAppearance = localAsset;
        }

		public void ExecuteCommand(int CommandID)
		{
			//this method was used to notify when an AddIn command was executed
			//the CommandID parameter identifies the command that was executed
    
			//Note:this method is now obsolete, you should use the new
			//ControlDefinition objects to implement commands, they have
			//their own event sinks to notify when the command is executed
		}

		public object Automation
		{
			//if you want to return an interface to another client of this addin,
			//implement that interface in a class and return that class object 
			//through this property

			get
			{
				return null;
			}
		}

		#endregion
	}
}
