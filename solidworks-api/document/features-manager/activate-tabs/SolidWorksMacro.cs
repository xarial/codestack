using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Forms;

using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;


namespace ActivateFeatMgrTab
{
    public partial class SolidWorksMacro
    {
        public void Main()
        {
            var model = swApp.IActiveDoc2;

            try
            {
                if (model != null)
                {
                    swApp.SendMsgToUser(string.Format("Active Feature Manager Tab: {0}", model.GetActiveStandardFeatureManagerTab()));

                    model.ActivateStandardFeatureManagerTab(FeatMgrTab_e.DisplayManager);
                }
                else
                {
                    throw new NullReferenceException("Model is not opened");
                }
            }
            catch(Exception ex)
            {
                swApp.SendMsgToUser2(ex.Message, (int)swMessageBoxIcon_e.swMbStop, (int)swMessageBoxBtn_e.swMbOk);
            }

            return;
        }

        public SldWorks swApp;

    }
}

