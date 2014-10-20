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
Public Class EventHandler_Initialisatie_OnEntry
    Inherits RoutingEventHandlerBase

    Dim HuidigeKlasse As String = Me.GetType.Name
    Dim LogFileNaam As String = ""
    Dim Params As Object = Arco.Doma.Library.Helpers.ArcoInfo.GetParameters("Digitaal Toezicht", "Digitaal Toezicht")

#Region " Execute "

    Public Overrides Sub Execute(ByVal WFCurrentCase As cCase)

        ' bepaal de naam van de logfile :
        Try
            LogFileNaam = Params.GetValue("Logging", "LogFileName", "")
            LogFileNaam = LogFileNaam.Replace("[ROOT]", Arco.Settings.FrameWorkSettings.GetGlobalVar("ROOT", True)).Replace("[DATE]", Now.ToString("yyyyMMdd"))
        Catch ex As Exception
            LogFileNaam = ""
        End Try

        Dim HuidigeMethod As String = Reflection.MethodInfo.GetCurrentMethod.Name
        Arco.Utils.Logging.Log("Start uitvoering : " & HuidigeKlasse & "-" & HuidigeMethod & " (case=" & WFCurrentCase.Case_ID & ")", LogFileNaam)

        ' dossier behoort tot de onderneming die gekozen werd (of automatisch geselecteerd werd) bij het aanloggen
        Dim lsondernemingsnummer As String
        lsondernemingsnummer = GetVariable("Ondernemingsnummer")
        WFCurrentCase.SetProperty("Ondernemingsnummer", lsondernemingsnummer)
        Arco.Utils.Logging.Log(vbTab & "Ondernemingsnummer = " & lsondernemingsnummer, LogFileNaam)

        WFCurrentCase.SetProperty("Bestuur", UCase(GetVariable("Bestuur")))
        WFCurrentCase.SetProperty("Type bestuur", GetVariable("Type bestuur"))
        WFCurrentCase.SetProperty("TypeOrgTypeBest", GetVariable("TypeOrgTypeBest"))

        'de thema's die de ingelogde user mag opladen/goedkeuren, zitten al in de sessie variabele themas
        WFCurrentCase.SetProperty("Themas_1", GetVariable("Themas_1"))
        Arco.Utils.Logging.Log("Einde uitvoering : " & HuidigeKlasse & "-" & HuidigeMethod, LogFileNaam)

    End Sub
    Private Function GetVariable(ByVal vsKey As String) As Object
        If Arco.Doma.Library.Security.BusinessIdentity.CurrentIdentity.Properties.ContainsKey(vsKey) Then
            Return Arco.Doma.Library.Security.BusinessIdentity.CurrentIdentity.Properties.Item(vsKey)
        Else
            Return Nothing
        End If
    End Function

#End Region




End Class
