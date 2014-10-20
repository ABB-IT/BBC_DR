' **************************************************************************
' Project naming convention : [Customer].DocRoom.[Name]
' SourceSafe location : $/Doma60/Custom/[Customer].DocRoom.[Name]
' **************************************************************************
' Author :
' Created by         on   
' Modified by        on 
' Description :
' **************************************************************************
Imports Arco.Doma.Library.baseObjects
Imports Arco.Doma.Library.Routing
Imports Arco.Utils.Logging
Public Class EventHandler_OpladenDocumenten_OnKeep
    Inherits RoutingEventHandlerBase

#Region " Execute "

    Public Overrides Sub Execute(ByVal WFCurrentCase As cCase)
        'Debug("EventHandler_OpladenDocumenten_OnKeep : Case " & WFCurrentCase.Tech_ID)
    End Sub

#End Region

End Class
