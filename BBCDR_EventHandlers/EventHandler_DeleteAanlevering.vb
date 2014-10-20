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
Public Class EventHandler_DeleteAanlevering
    Inherits RoutingEventHandlerBase

#Region " Execute "

    Public Overrides Sub Execute(ByVal WFCurrentCase As cCase)
        'Debug("EventHandler_DeleteAanlevering : Case " & WFCurrentCase.Tech_ID)
        ' **************************************************************************
        ' Project naming convention : [Customer].DocRoom.[Name]
        ' SourceSafe location : $/Doma60/Custom/[Customer].DocRoom.[Name]
        ' **************************************************************************
        ' Author :
        ' Created by         on   
        ' Modified by        on 
        ' Description :
        ' **************************************************************************

        WFCurrentCase.DispatchStop(False, True)

    
    End Sub

#End Region

End Class
