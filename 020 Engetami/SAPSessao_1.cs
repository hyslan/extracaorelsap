using SAPFEWSELib;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SapSessao
{
    public class SapSessao1
    {
        public GuiSession Sessao()
        {
            //Get the Windows Running Object Table
            SapROTWr.CSapROTWrapper sapROTWrapper = new SapROTWr.CSapROTWrapper();
            //Get the ROT Entry for the SAP Gui to connect to the COM
            object SapGuilRot = sapROTWrapper.GetROTEntry("SAPGUI");
            //Get the reference to the Scripting Engine
            object engine = SapGuilRot.GetType().InvokeMember("GetScriptingEngine", System.Reflection.BindingFlags.InvokeMethod, null, SapGuilRot, null);
            //Get the reference to the running SAP Application Window
            GuiApplication GuiApp = (GuiApplication)engine;
            //Get the reference to the first open connection
            GuiConnection connection = (GuiConnection)GuiApp.Connections.ElementAt(0);
            //get the first available session
            GuiSession session = (GuiSession)connection.Children.ElementAt(1);
            //Get the reference to the main "Frame" in which to send virtual key commands
            GuiFrameWindow frame = (GuiFrameWindow)session.FindById("wnd[0]");

            return session;
        }
    }
}
