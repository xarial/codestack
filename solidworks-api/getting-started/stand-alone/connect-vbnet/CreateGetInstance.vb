Const PROG_ID As String = "SldWorks.Application"

'Using Interaction.CreateObject function
Dim app1 = TryCast(CreateObject(PROG_ID), SolidWorks.Interop.sldworks.ISldWorks)
app1.Visible = True

'Using Interaction.GetObject function
Dim app2 = TryCast(GetObject("", PROG_ID), SolidWorks.Interop.sldworks.ISldWorks)
app2.Visible = True

'Using Activator
Dim progType = System.Type.GetTypeFromProgID(PROG_ID)
Dim app3 = TryCast(System.Activator.CreateInstance(progType), SolidWorks.Interop.sldworks.ISldWorks)
app3.Visible = True
