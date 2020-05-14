Imports System.IO
Imports System.Runtime.InteropServices
Imports System.Runtime.InteropServices.ComTypes

Public Class ComStream
    Inherits Stream

    Private ReadOnly m_ComStream As IStream
    Private ReadOnly m_Commit As Boolean
    Private m_IsWritable As Boolean

    Public Sub New(ByRef comStream As IStream, writable As Boolean, Optional commit As Boolean = True)

        If comStream Is Nothing Then
            Throw New ArgumentNullException(NameOf(comStream))
        End If

        m_ComStream = comStream
        m_IsWritable = writable
        m_Commit = commit

    End Sub

    Public Overrides ReadOnly Property CanRead() As Boolean
        Get
            Return True
        End Get
    End Property

    Public Overrides ReadOnly Property CanSeek() As Boolean
        Get
            Return True
        End Get
    End Property

    Public Overrides ReadOnly Property CanWrite() As Boolean
        Get
            Return m_IsWritable
        End Get
    End Property

    Public Overrides ReadOnly Property Length As Long
        Get
            Const STATSFLAG_NONAME As Integer = 1

            Dim stats As ComTypes.STATSTG = Nothing
            m_ComStream.Stat(stats, STATSFLAG_NONAME)

            Return stats.cbSize
        End Get
    End Property

    Public Overrides Property Position() As Long
        Get
            Return Seek(0, SeekOrigin.Current)
        End Get
        Set(ByVal Value As Long)
            Seek(Value, SeekOrigin.Begin)
        End Set
    End Property

    Public Overrides Sub Flush()
        If m_Commit Then
            Const STGC_DEFAULT As Integer = 0
            m_ComStream.Commit(STGC_DEFAULT)
        End If
    End Sub

    Public Overrides Sub SetLength(ByVal Value As Long)
        m_ComStream.SetSize(Value)
    End Sub

    Public Overrides Sub Write(buffer() As Byte, offset As Integer, count As Integer)
        If offset <> 0 Then
            Dim bufferSize As Integer
            bufferSize = buffer.Length - offset
            Dim tmpBuffer(bufferSize) As Byte
            Array.Copy(buffer, offset, tmpBuffer, 0, bufferSize)
            m_ComStream.Write(tmpBuffer, bufferSize, Nothing)
        Else
            m_ComStream.Write(buffer, count, Nothing)
        End If
    End Sub

    Public Overrides Function Read(buffer() As Byte, offset As Integer, count As Integer) As Integer

        Dim bytesRead As Integer = 0
        Dim boxBytesRead As Object = bytesRead
        Dim hObject As GCHandle

        Try
            hObject = GCHandle.Alloc(boxBytesRead, GCHandleType.Pinned)
            Dim pBytesRead As IntPtr = hObject.AddrOfPinnedObject()

            If offset <> 0 Then
                Dim tmpBuffer(count - 1) As Byte
                m_ComStream.Read(tmpBuffer, count, pBytesRead)
                bytesRead = CInt(boxBytesRead)
                Array.Copy(tmpBuffer, 0, buffer, offset, bytesRead)
            Else
                m_ComStream.Read(buffer, count, pBytesRead)
                bytesRead = CInt(boxBytesRead)
            End If

        Finally
            If hObject.IsAllocated Then
                hObject.Free()
            End If
        End Try

        Return bytesRead

    End Function

    Public Overrides Function Seek(offset As Long, origin As SeekOrigin) As Long

        Dim curPosition As Long = 0
        Dim boxCurPosition As Object = curPosition
        Dim hObject As GCHandle

        Try
            hObject = GCHandle.Alloc(boxCurPosition, GCHandleType.Pinned)
            Dim pCurPosition As IntPtr = hObject.AddrOfPinnedObject()

            m_ComStream.Seek(offset, origin, pCurPosition)
            curPosition = CLng(boxCurPosition)
        Finally
            If hObject.IsAllocated Then
                hObject.Free()
            End If
        End Try

        Return curPosition
    End Function

    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing Then
                m_IsWritable = False
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    Protected Overrides Sub Finalize()
        Dispose(False)
    End Sub

End Class
