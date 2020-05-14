using EdmLib;
using System;
using System.Runtime.InteropServices;

namespace CodeStack
{
    [ComVisible(true)]
    [Guid("3A601AFC-7007-46A7-9E71-D3BD41B5E2E2")]
    public class PdmAddInSample : IEdmAddIn5
    {
        const int TEST_CMD_ID = 1;

        public void GetAddInInfo(ref EdmAddInInfo poInfo, IEdmVault5 poVault, IEdmCmdMgr5 poCmdMgr)
        {
            poInfo.mbsAddInName = "Demo AddIn";
            poInfo.mlRequiredVersionMajor = 17; //SOLIDWORKS PDM 2017 SP0

            poCmdMgr.AddCmd(TEST_CMD_ID, "Test Menu Command");
        }

        public void OnCmd(ref EdmCmd poCmd, ref Array ppoData)
        {
            if (poCmd.meCmdType == EdmCmdType.EdmCmd_Menu)
            {
                if (poCmd.mlCmdID == TEST_CMD_ID)
                {
                    (poCmd.mpoVault as IEdmVault10).MsgBox(0, "Hello World!");
                }
            }
        }
    }
}
