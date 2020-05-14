Imports System.Runtime.InteropServices

Public Class OpenGl

	<DllImport("opengl32")>
	Public Shared Sub glBegin(ByVal mode As UInteger)
	End Sub

	<DllImport("opengl32")>
	Public Shared Sub glEnd()
	End Sub

	<DllImport("opengl32")>
	Public Shared Sub glVertex3d(ByVal x As Double, ByVal y As Double, ByVal z As Double)
	End Sub

	<DllImport("opengl32.dll")>
	Public Shared Sub glDisable(ByVal cap As UInteger)
	End Sub

	<DllImport("opengl32.dll")>
	Public Shared Sub glColor4f(ByVal R As Single, ByVal G As Single, ByVal B As Single, ByVal A As Single)
	End Sub

	<DllImport("opengl32.dll")>
	Public Shared Sub glEnable(ByVal cap As UInteger)
	End Sub

	<DllImport("opengl32.dll")>
	Public Shared Sub glPolygonMode(ByVal face As UInteger, ByVal mode As UInteger)
	End Sub

	<DllImport("opengl32.dll")>
	Public Shared Sub glLineWidth(ByVal width As Single)
	End Sub

	<DllImport("opengl32.dll")>
	Public Shared Sub glLineStipple(ByVal factor As Integer, ByVal pattern As UShort)
	End Sub

	Public Const GL_FRONT_AND_BACK As Integer = &H408
	Public Const GL_LINE As UInteger = &H1B01
	Public Const GL_FILL As UInteger = &H1B02

	Public Const GL_TRIANGLES As UInteger = &H4
	Public Const GL_LINE_LOOP As UInteger = &H2
	Public Const GL_LIGHTING As UInteger = &HB50
	Public Const GL_LINE_SMOOTH As UInteger = &HB20
	Public Const GL_LINE_STIPPLE As UInteger = &HB24

End Class
