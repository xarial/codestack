Const PROG_ID As String = "SldWorks.Application"
Dim app = TryCast(System.Runtime.InteropServices.Marshal.GetActiveObject(PROG_ID),
	SolidWorks.Interop.sldworks.ISldWorks)