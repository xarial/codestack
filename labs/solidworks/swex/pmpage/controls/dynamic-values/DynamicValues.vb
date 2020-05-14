Imports System.ComponentModel

Public Class DynamicValuesDataModel
    Implements INotifyPropertyChanged

    Public Event PropertyChanged As PropertyChangedEventHandler Implements INotifyPropertyChanged.PropertyChanged

    Private m_Val1 As Double
    Private m_Val2 As Double

    Public Property Val1 As Double
        Get
            Return m_Val1
        End Get
        Set(ByVal value As Double)
            m_Val1 = value
            RaiseEvent PropertyChanged(Me, New PropertyChangedEventArgs(NameOf(Val1)))
            Val2 = value * 2
        End Set
    End Property

    Public Property Val2 As Double
        Get
            Return m_Val2
        End Get
        Set(ByVal value As Double)
            m_Val2 = value
            RaiseEvent PropertyChanged(Me, New PropertyChangedEventArgs(NameOf(Val2)))
        End Set
    End Property

    Public ReadOnly Property Reset As Action
        Get
            Return AddressOf OnResetClick
        End Get
    End Property

    Private Sub OnResetClick()
        Val1 = 10
    End Sub

End Class
