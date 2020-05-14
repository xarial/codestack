Imports CodeStack.SwEx.AddIn.Base
Imports SolidWorks.Interop.sldworks
Imports System.Drawing
Imports SolidWorks.Interop.swconst
Imports CodeStack.OglTetrahedron.OpenGl

Public Class OpenGlDocumentHandler
	Implements IDocumentHandler

	ReadOnly m_FaceColor As Color = Color.Green
	ReadOnly m_EdgeColor As Color = Color.Black

	Dim m_MathUtils As IMathUtility
	Dim m_Model As IModelDoc2
	Dim m_View As ModelView

	Public Sub Init(ByVal app As ISldWorks, ByVal model As IModelDoc2) Implements IDocumentHandler.Init

		m_MathUtils = app.IGetMathUtility
		m_Model = model
		m_View = model.IActiveView

		If m_View IsNot Nothing Then
			AddHandler m_View.BufferSwapNotify, AddressOf OnBufferSwapNotify
		End If

	End Sub

	Private Function OnBufferSwapNotify() As Integer

		Dim a As Double() = New Double() {0, 0, 0}
		Dim b As Double() = New Double() {1, 0, 0}
		Dim c As Double() = New Double() {0.5, Math.Sqrt(3) / 2, 0}
		Dim d As Double() = New Double() {0.5, Math.Sqrt(3) / 6, Math.Sqrt(6) / 3}

		Select Case CType(m_View.DisplayMode, swViewDisplayMode_e)
			Case swViewDisplayMode_e.swViewDisplayMode_ShadedWithEdges
				DrawTetrahedron(m_FaceColor, True, False, False, 3.0F)
				DrawTetrahedron(m_EdgeColor, True, True, False, 3.0F)
			Case swViewDisplayMode_e.swViewDisplayMode_Shaded
				DrawTetrahedron(m_FaceColor, True, False, False, 3.0F)
			Case swViewDisplayMode_e.swViewDisplayMode_HiddenLinesRemoved '
				DrawTetrahedron(m_EdgeColor, False, False, False, 3.0F)
			Case swViewDisplayMode_e.swViewDisplayMode_HiddenLinesGrayed '
				DrawTetrahedron(m_EdgeColor, True, True, True, 1.0F)
			Case swViewDisplayMode_e.swViewDisplayMode_Wireframe
				DrawTetrahedron(m_EdgeColor, True, True, False, 3.0F)
		End Select

		Dim pt1 As IMathPoint = m_MathUtils.CreatePoint(New Double() {0, 0, 0})
		Dim pt2 As IMathPoint = m_MathUtils.CreatePoint(New Double() {1, 1, 1})

		m_Model.Extension.SetVisibleBox(pt1, pt2)

		Return 0

	End Function

	Private Sub DrawTetrahedron(color As Color, fill As Boolean, wireframe As Boolean, dashed As Boolean, width As Single)

		Dim a As Double() = New Double() {0, 0, 0}
		Dim b As Double() = New Double() {1, 0, 0}
		Dim c As Double() = New Double() {0.5, Math.Sqrt(3) / 2, 0}
		Dim d As Double() = New Double() {0.5, Math.Sqrt(3) / 6, Math.Sqrt(6) / 3}

		DrawTriangle(a, c, b, color, fill, wireframe, dashed, width)
		DrawTriangle(a, d, c, color, fill, wireframe, dashed, width)
		DrawTriangle(c, d, b, color, fill, wireframe, dashed, width)
		DrawTriangle(d, a, b, color, fill, wireframe, dashed, width)

	End Sub

	Private Sub DrawTriangle(a() As Double, b() As Double, c() As Double, color As Color, fill As Boolean, wireframe As Boolean, dashed As Boolean, width As Single)

		glPolygonMode(GL_FRONT_AND_BACK, IIf(fill, GL_FILL, GL_LINE))

		glDisable(GL_LIGHTING)

		If wireframe Then

			glEnable(GL_LINE_SMOOTH)

			If dashed Then
				glEnable(GL_LINE_STIPPLE)
				glLineStipple(4, &HAAAA)
			End If

		End If

		glBegin(IIf(wireframe, GL_LINE_LOOP, GL_TRIANGLES))

		If wireframe Then
			glLineWidth(width)
		End If

		glColor4f(color.R / 255.0F, color.G / 255.0F, color.B / 255.0F, color.A / 255.0F)
		glVertex3d(a(0), a(1), a(2))
		glVertex3d(b(0), b(1), b(2))
		glVertex3d(c(0), c(1), c(2))

		glEnd()

		glDisable(GL_LINE_SMOOTH)
		glDisable(GL_LINE_STIPPLE)

	End Sub

	Public Sub Dispose() Implements IDisposable.Dispose
		If m_View IsNot Nothing Then
			RemoveHandler m_View.BufferSwapNotify, AddressOf OnBufferSwapNotify
		End If
	End Sub

End Class
