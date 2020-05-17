Public Module M_PublicVars

    'Declare Vars
    Public pv_dbo_sqlConnectionString As String
    Public pv_env_showErrors As Boolean

    'Declare Singletons
    Public ps_registry As S_Registry = S_Registry.GetInstance()

End Module