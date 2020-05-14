Imports System.Runtime.InteropServices
Imports EdmLib

Namespace CodeStack

    <ComVisible(True)>
    <Guid("1B2547B5-F1F9-457D-941C-D25668619101")>
    Public Class SwPdmAddInVB
        Implements IEdmAddIn5

        Const CMD_ID As Integer = 1

        Public Sub GetAddInInfo(ByRef poInfo As EdmAddInInfo, poVault As IEdmVault5, poCmdMgr As IEdmCmdMgr5) _
        Implements IEdmAddIn5.GetAddInInfo

            poInfo.mbsAddInName = "Demo AddIn"
            poInfo.mlRequiredVersionMajor = 17

            poCmdMgr.AddCmd(CMD_ID, "Test Menu Command")

        End Sub

        Public Sub OnCmd(ByRef poCmd As EdmCmd, ByRef ppoData As Array) Implements IEdmAddIn5.OnCmd

            If poCmd.meCmdType = EdmCmdType.EdmCmd_Menu Then

                If poCmd.mlCmdID = CMD_ID Then
                    CType(poCmd.mpoVault, IEdmVault10).MsgBox(0, "Hello World!")

                End If

            End If

        End Sub
    End Class

End Namespace
