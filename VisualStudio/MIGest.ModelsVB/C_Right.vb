Public Class Right

    'Instanciation Member
    Private _label As String
    Private _descr As String
    Private _factive As Boolean
    Private _comm As String

    'Constructor
    Public Sub New(label As String, descr As String, factive As Boolean, Optional comm As String = vbNullString)
        Me.Label = label
        Me.Descr = descr
        Me.IsActive = factive
        Me.Comm = comm
    End Sub

    Public Sub New(usn)
        '-- A faire avec la DB --
    End Sub


    'Getter / Setter
    Public Property Label As String
        Get
            Return _label
        End Get
        Set
            Try
                If Value.Trim().Length > 0 Then
                    _label = Value
                Else
                    Throw New Exception("The value cannot be empty")
                End If
            Catch ex As Exception
                Throw
            End Try
        End Set
    End Property

    Public Property Descr As String
        Get
            Return _descr
        End Get
        Set
            Try
                If Value.Trim().Length > 0 Then
                    _descr = Value
                Else
                    Throw New Exception("The value cannot be empty")
                End If
            Catch ex As Exception
                Throw
            End Try
        End Set
    End Property

    Public Property IsActive As Boolean
        Get
            Return _factive
        End Get
        Set
            _factive = Value
        End Set
    End Property

    Public Property Comm As String
        Get
            Return _comm
        End Get
        Set
            Try
                If Value.Trim().Length > 0 Then
                    _comm = Value
                Else
                    Throw New Exception("The value cannot be empty")
                End If
            Catch ex As Exception
                Throw
            End Try
        End Set
    End Property

    'Implement Static Methods
    Private Shared Function _ListOfElements() As Right()
        '-- A Faire avec la DB --
    End Function

    'Implement Instance Methods


    'Implement Overrides


    'Interface Methods
    Public Shared Function ListOfElements() As Right()
        Return _ListOfElements()
    End Function




End Class