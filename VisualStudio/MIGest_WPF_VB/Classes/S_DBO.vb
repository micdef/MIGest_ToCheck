'################################################################################
'# Module Name           : S_DBO                                                #
'# Module Description    : Singleton class what contains the database functions #
'# Module Creation Date  : 04/14/2020                                           #
'# Module Creator        : Defraene Michaël (AKA Sonic)                         #
'# Module Licensed to    : Defraene Michaël (AKA Sonic)                         #
'# Copyrights            : Defraene Michaël (AKA Sonic)                         #
'# Version               : 1.0                                                  #
'################################################################################

'********************************************************************************
'* PART 0 : Imports                                                             *
'********************************************************************************
Imports System.IO
Imports System.Configuration
Imports System.Data.Odbc
Imports MIGest_WPF_VB.S_Registry

Public Class S_DBO

    '********************************************************************************
    '* PART 1 : Static Members                                                      *
    '********************************************************************************
    Private Shared pathODBC As String = "SOFTWARE\ODBC\ODBC.INI"
    Private Shared _instance As S_DBO = Nothing
    Private Shared reg As S_Registry = S_Registry.GetInstance()

    '********************************************************************************
    '* PART 2 : Class Members                                                       *
    '********************************************************************************
    Private dsn As String
    Private odbcConn As OdbcConnection
    Private odbcComm As OdbcCommand
    Private odbcDataAdap As OdbcDataAdapter
    Private odbcState As Boolean

    '********************************************************************************
    '* PART 3 : Constructors                                                        *
    '********************************************************************************
    Private Sub New()
        Try

        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    '********************************************************************************
    '* PART 4 : Singleton Instance                                                  *
    '********************************************************************************

    '********************************************************************************
    '* PART 5 : Getters                                                             *
    '********************************************************************************

    '********************************************************************************
    '* PART 6 : Setters                                                             *
    '********************************************************************************

    '********************************************************************************
    '* PART 7 : Class functions                                                     *
    '********************************************************************************

    '********************************************************************************
    '* PART 8 : Class Subs                                                          *
    '********************************************************************************
    Private Sub changeConnection(ByVal dsn As String)
        My.Settings.Item("MyConnectionString") = "Dsn=" & dsn
    End Sub

    '********************************************************************************
    '* PART 9 : Static Functions                                                    *
    '********************************************************************************
    Public Shared Function listDSN() As Object
        Dim lstDsn As New List(Of String)
        Dim sr As StreamReader
        Dim dsn As String
        Try
            If My.Computer.Registry.CurrentUser.OpenSubKey(pathODBC).SubKeyCount > 0 Then
                For Each dsn In My.Computer.Registry.CurrentUser.OpenSubKey(pathODBC).GetSubKeyNames()
                    lstDsn.Add(dsn)
                Next
            ElseIf My.Computer.Registry.LocalMachine.OpenSubKey(pathODBC).SubKeyCount > 0 Then
                For Each dsn In My.Computer.Registry.LocalMachine.OpenSubKey(pathODBC).GetSubKeyNames()
                    lstDsn.Add(dsn)
                Next
            Else
                sr = New StreamReader("C:\Windows\ODBC.INI")
                Do
                    dsn = sr.ReadLine()
                    If (Not dsn Is Nothing) Then
                        lstDsn.Add(dsn)
                    End If
                Loop While (Not dsn Is Nothing)
            End If
            Return lstDSN
        Catch ex As Exception
            Throw ex
        End Try
    End Function

    '********************************************************************************
    '* PART 10 : Static Subs                                                        *
    '********************************************************************************

    '********************************************************************************
    '* PART 11 : Enumerations                                                       *
    '********************************************************************************


End Class