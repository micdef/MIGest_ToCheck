'################################################################################
'# Module Name           : S_Registry                                           #
'# Module Description    : Singleton class what contains the registry functions #
'# Module Creation Date  : 04/10/2020                                           #
'# Module Creator        : Defraene Michaël (AKA Sonic)                         #
'# Module Licensed to    : Defraene Michaël (AKA Sonic)                         #
'# Copyrights            : Defraene Michaël (AKA Sonic)                         #
'# Version               : 1.0                                                  #
'################################################################################

'********************************************************************************
'* PART 0 : Imports                                                             *
'********************************************************************************
' N/A

Public Class S_Registry

    '********************************************************************************
    '* PART 1 : Static Members                                                      *
    '********************************************************************************
    Private Shared _instance As S_Registry = Nothing

    '********************************************************************************
    '* PART 2 : Class Members                                                       *
    '********************************************************************************
    Private pathKey As String = My.Computer.Registry.CurrentUser.ToString() & "\SOFTWARE\SOCO\MIGest"

    '********************************************************************************
    '* PART 3 : Constructors                                                        *
    '********************************************************************************
    Private Sub New() 'Empty Constructor
    End Sub

    '********************************************************************************
    '* PART 4 : Singleton Instance                                                  *
    '********************************************************************************
    ''' <summary>
    '''     Function what give or generate the instance of the class. Only 1 instance
    '''     is possible.
    ''' </summary>
    ''' <returns>{S_Registry} The instance of the class</returns>
    Public Shared Function GetInstance() As S_Registry
        Try
            If _instance Is Nothing Then
                _instance = New S_Registry()
            End If
            Return _instance
        Catch ex As Exception
            Throw ex
        End Try
    End Function

    '********************************************************************************
    '* PART 5 : Getters                                                             *
    '********************************************************************************
    ' N/A

    '********************************************************************************
    '* PART 6 : Setters                                                             *
    '********************************************************************************
    ' N/A

    '********************************************************************************
    '* PART 7 : Class functions                                                     *
    '********************************************************************************
    ''' <summary>
    '''     Function what gets the key value
    ''' </summary>
    ''' <param name="myKey">[Required] {String} The name of the key</param>
    ''' <param name="myFolder">[Optional] {String} The folder path of the key</param>
    ''' <returns>{String} The value what the key contains</returns>
    Public Function GetRegKey(ByVal myKey As String, Optional ByVal myFolder As String = vbNullString) As String
        Dim myValue As String
        Try
            myValue = My.Computer.Registry.GetValue(pathKey & IIf(myFolder = vbNullString, "", "\" & myFolder), myKey, Nothing)
            Return myValue
        Catch ex As Exception
            If pv_env_showErrors Then
                MsgBox("Une erreur est survenue : " & vbNewLine & vbNewLine &
                       ex.Message & vbNewLine & vbNewLine &
                       "Veuillez réessayer. Si l'erreur persiste, veuillez contacter votre administrateur.",
                       MsgBoxStyle.Critical, My.Application.Info.ProductName & " V:" & My.Application.Info.Version.ToString())
            End If
            Return Nothing
        End Try
    End Function

    '********************************************************************************
    '* PART 8 : Class Subs                                                          *
    '********************************************************************************
    ''' <summary>
    '''     Function what create the key in the registry
    ''' </summary>
    ''' <param name="myKey">[Required] {String} The name of the key</param>
    ''' <param name="myFolder">[Optional] {String} The folder path of the key</param>
    Public Sub CreateRegKey(ByVal myKey As String, Optional ByVal myFolder As String = vbNullString)
        Try
            My.Computer.Registry.SetValue(pathKey & "\" & IIf(myFolder = vbNullString, "", myFolder & "\"), myKey, Microsoft.Win32.RegistryValueKind.String)
        Catch ex As Exception
            If pv_env_showErrors Then
                MsgBox("Une erreur est survenue : " & vbNewLine & vbNewLine &
                       ex.Message & vbNewLine & vbNewLine &
                       "Veuillez réessayer. Si l'erreur persiste, veuillez contacter votre administrateur.",
                       MsgBoxStyle.Critical, My.Application.Info.ProductName & " V:" & My.Application.Info.Version.ToString())
            End If
        End Try
    End Sub

    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="myKey">[Required] {String} The name of the key</param>
    ''' <param name="myValue">[Required] {String} The value of the key</param>
    ''' <param name="myFOlder">[Optional] {String} The folder path of the key</param>
    Public Sub SetRegKey(ByVal myKey As String, ByVal myValue As String, Optional ByVal myFOlder As String = vbNullString)
        Try
            My.Computer.Registry.SetValue(pathKey & "\" & IIf(myFOlder = vbNullString, "", myFOlder & "\"), myKey, myValue)
        Catch ex As Exception
            If pv_env_showErrors Then
                MsgBox("Une erreur est survenue : " & vbNewLine & vbNewLine &
                       ex.Message & vbNewLine & vbNewLine &
                       "Veuillez réessayer. Si l'erreur persiste, veuillez contacter votre administrateur.",
                       MsgBoxStyle.Critical, My.Application.Info.ProductName & " V:" & My.Application.Info.Version.ToString())
            End If
        End Try
    End Sub

    '********************************************************************************
    '* PART 9 : Static Functions                                                    *
    '********************************************************************************
    ' N/A

    '********************************************************************************
    '* PART 10 : Static Subs                                                        *
    '********************************************************************************
    ' N/A

    '********************************************************************************
    '* PART 11 : Enumerations                                                       *
    '********************************************************************************
    ' N/A

End Class