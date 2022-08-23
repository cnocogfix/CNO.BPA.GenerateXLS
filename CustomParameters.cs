using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Drawing.Design;
using System.ComponentModel;
using Emc.InputAccel.QuickModule.ClientScriptingInterface;
//using Emc.InputAccel.ScanPlus.Scripting;

namespace CNO.BPA.ScanPop
{
    [Serializable(),
    CustomParameterEditor(typeof(CustomParameterEditorController))]
    public class CustomParameters
    {
        string _username = string.Empty;
        string _password = string.Empty;
        string _dsn = string.Empty;

        public string UserName
        {
            get { return _username; }
            set { _username = value; }
        }
        public string Password
        {
            get { return _password; }
            set { _password = value; }
        }
        public string DSN
        {
            get { return _dsn; }
            set { _dsn = value; }
        }
    }
}