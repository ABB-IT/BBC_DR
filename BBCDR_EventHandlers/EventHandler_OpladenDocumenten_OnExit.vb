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
Imports Ionic.Zip
Imports System.IO
Imports System.Xml
Imports System.Xml.Serialization
Imports System.Xml.Schema


Public Class EventHandler_OpladenDocumenten_OnExit
    Inherits RoutingEventHandlerBase

    Dim HuidigeKlasse As String = Me.GetType.Name
    Dim LogFileNaam As String = ""
    Dim bEndedSuccessFully As Boolean = True

#Region " Execute "

    Public Overrides Sub Execute(ByVal WFCurrentCase As cCase)

        Arco.Security.BusinessIdentity.CurrentIdentity.RunElevated()
        Dim Params As Arco.Doma.Library.Helpers.ArcoInfo = Arco.Doma.Library.Helpers.ArcoInfo.GetParameters("Digitaal Toezicht", "Digitaal Toezicht")
        ' bepaal de naam van de logfile :
        Try
            LogFileNaam = Params.GetValue("Logging", "LogFileName", "")
            LogFileNaam = LogFileNaam.Replace("[ROOT]", Arco.Settings.FrameWorkSettings.GetGlobalVar("ROOT", True)).Replace("[DATE]", Now.ToString("yyyyMMdd"))
        Catch ex As Exception
            LogFileNaam = ""
        End Try

        Dim HuidigeMethod As String = Reflection.MethodInfo.GetCurrentMethod.Name
        Arco.Utils.Logging.Log("Start uitvoering : " & HuidigeKlasse & "-" & HuidigeMethod & " (case=" & WFCurrentCase.Case_ID & ")", LogFileNaam)
        If CheckDocumentIsValid(WFCurrentCase) = True Then
            Call CreateZipFile(WFCurrentCase, Params)
        End If
        Arco.Utils.Logging.Log("Einde uitvoering : " & IIf(bEndedSuccessFully, " succes.", " met fouten!" & HuidigeKlasse & "-" & HuidigeMethod), LogFileNaam)

    End Sub

    Private Function CheckDocumentIsValid(ByVal WFCurrentCase As Arco.Doma.Library.Routing.cCase) As Boolean

        'Functioneel nut :
        ' Controleer de bestanden. Als er 2 bestanden aanwezig zijn (één met extensie 'XML' en één met extensie 'XBRL' dan is de upload OK
        ' Indien niet wordt de case gereject.
        '
        'Parameters : 
        '   WFCurrentCase : de Routing workflow procedure van waaruit de code wordt getriggerd.

        Dim HuidigeMethod As String = Reflection.MethodInfo.GetCurrentMethod.Name
        Arco.Utils.Logging.Log(vbTab & HuidigeMethod & " gestart.", LogFileNaam)

        Dim lodocument As Arco.Doma.Library.Document
        lodocument = WFCurrentCase.TargetObject

        Dim lbupload_ok As Boolean = True
        Dim lixmlcount As Integer = 0
        Dim lixbrlcount As Integer = 0
        Dim lsxmlfile As String = ""
        Dim lsxbrlfile As String = ""

        If lodocument.FileCount <> 2 Then lbupload_ok = False
        For Each lofile As Arco.Doma.Library.FileList.FileInfo In lodocument.Files
            Dim loFsFile As Arco.Doma.Library.FileServers.DM_FileServer = Arco.Doma.Library.FileServers.DM_FileServer.GetFileServerByID(lofile.FILESERVER_ID)
            Dim lsfilename As String = ""
            lsfilename = lofile.FILE_PATH
            Arco.Utils.Logging.Log(vbTab & "File = " & loFsFile.RealBasePath & lofile.FILE_PATH, LogFileNaam)

            If InStr(lsfilename, ".xml") Then
                lixmlcount += 1
                lsxmlfile = loFsFile.RealBasePath & lofile.FILE_PATH
            ElseIf InStr(lsfilename, ".xbrl") Then
                lixbrlcount += 1
                lsxbrlfile = loFsFile.RealBasePath & lofile.FILE_PATH
            End If
        Next

        If Not (lbupload_ok And lixmlcount = 1 And lixbrlcount = 1) Then
            WFCurrentCase.RejectComment = "Gelieve twee documenten toe te voegen aub. Eén met extensie 'xml' en één met extensie 'xbrl'."
            WFCurrentCase.RejectUser = "BBCDR"
            WFCurrentCase.SetProperty("Rejected", True)
            Arco.Utils.Logging.Log(vbTab & "Het aantal bestanden voldoet niet aan de vereisten: Eén met extensie 'xml' en één met extensie 'xbrl'", LogFileNaam)
            Arco.Utils.Logging.Log(vbTab & "Rejected = True", LogFileNaam)
            'WFCurrentCase.Save()
            'Exit Function
            Return False
        Else
            WFCurrentCase.SetProperty("Rejected", False)
            Arco.Utils.Logging.Log(vbTab & "Rejected = False", LogFileNaam)
            'getdatafromfile wordt niet uitgevoerd omwille van beperking tot taxonomie 01.00
            'getdatafromfile(WFCurrentCase, lsxmlfile)
        End If

        'check if goedkeuring vereist of niet
        'check in de tabel PROJBEST_PROJTHEM_PROJAUTH_GK of de aanlevering dient gevalideerd te worden
        'indien er geen record gevonden wordt in projbest_projthem_projauth_gk is er geen validatie vereist

        ' DBe 2014/10/13: er wordt geen goedkeuring meer vereist. Het ophalen van de parameter vanuit de database is niet meer nodig.
        '                 ***********************************************************************************************************
        '' ''Dim loQueryABB As Arco.Server.DataQuery
        '' ''Dim loReaderABB As Arco.Server.SafeDataReader
        '' ''loQueryABB = New Arco.Server.DataQuery
        '' ''loQueryABB.ConnectionString = lsconnectionstring
        '' ''loQueryABB.Connect()
        '' ''loQueryABB.Clear()

        '' ''loQueryABB.Query = "SELECT  a.code FROM projbest_projthem_projauth_gk ppp,project_thema_rel pt, authorisatie a ,project_authorisatie_rel pa" & _
        '' ''    " WHERE ppp.projectbestuurid = " & GetVariable("BestuurProjectID_1") & " And pt.id = ppp.projectthemaid And pt.themaid in (0" & GetVariable("Themas_1") & ") And pt.projectid = 1" & _
        '' ''    " AND  ppp.projectauthorisatieid=pa.id AND pa.authorisatieid=a.id"
        '' ''Arco.Utils.Logging.Log(loQueryABB.Query)
        '' ''loReaderABB = loQueryABB.ExecuteReader()
        '' ''If loReaderABB.Read() Then
        '' ''    WFCurrentCase.SetProperty("Administratieve goedkeuring vereist", True)
        '' ''    Arco.Utils.Logging.Log("Goedkeuring vereist")
        '' ''Else
        'geen goedkeuring vereist
        WFCurrentCase.SetProperty("Administratieve goedkeuring vereist", False)
        Arco.Utils.Logging.Log(vbTab & "Geen goedkeuring vereist", LogFileNaam)
        '' ''End If
        '' ''loReaderABB.Close()

        '' ''If Not loQueryABB Is Nothing Then
        '' ''    loQueryABB.DisConnect()
        '' ''    loQueryABB = Nothing
        '' ''End If
        '' ''loQueryABB = Nothing
        '' ''loReaderABB = Nothing

        Arco.Utils.Logging.Log(vbTab & HuidigeMethod & " beëindigd.", LogFileNaam)
        Return True

    End Function

    Private Sub CreateZipFile(ByVal WFCurrentCase As Arco.Doma.Library.Routing.cCase, ByVal Params As Arco.Doma.Library.Helpers.ArcoInfo)
        'Functioneel nut :
        ' Het creëren van een zipfile en het resultaat ervan afleggen op de MFT\AanVlaanderen.
        '
        'Parameters : 
        '   WFCurrentCase : de Routing workflow procedure van waaruit de code wordt getriggerd.
        '   Params        : de lijst van ArcoInfo parameters. Deze werd in de hoofd method Execute ingelezen. Ze wordt doorgegeven om het inlezen van 
        '                   dezelfde lijst slechts één maal te moeten uitvoeren.

        Dim datetimeexport As String = Now.ToString("yyyyMMddhhmmssfff")
        Dim tempdir As String = String.Empty
        Dim mftdir As String = String.Empty
        Dim HuidigeMethod As String = Reflection.MethodInfo.GetCurrentMethod.Name

        Try
            Arco.Utils.Logging.Log(vbTab & HuidigeMethod & " gestart.", LogFileNaam)
            'Arco.Utils.Logging.Log(vbTab & "Ophalen van het MFT pad" , LogFileNaam) '

            'Dim loSettings As Arco.Doma.Library.Helpers.ArcoInfo = Arco.Doma.Library.Helpers.ArcoInfo.GetParameters("Digitaal Toezicht", "Digitaal Toezicht")
            mftdir = Params.GetValue("Paths", "Path MFT", "")
            If Not mftdir.EndsWith("\") Then mftdir &= "\"
            Arco.Utils.Logging.Log(vbTab & "Opgehaald MFT pad = " & mftdir, LogFileNaam)

            'Arco.Utils.Logging.Log(vbTab & "Ophalen MFT Temp DT pad", LogFilenaam)
            tempdir = Params.GetValue("Paths", "Path MFT Temp BBC", "")
            If Not tempdir.EndsWith("\") Then tempdir &= "\"
            If Not Directory.Exists(tempdir & datetimeexport) Then
                Directory.CreateDirectory(tempdir & datetimeexport)
                Arco.Utils.Logging.Log(vbTab & "Tempdir " & tempdir & datetimeexport & " is gecreëerd.", LogFileNaam)
            End If
            tempdir = tempdir & datetimeexport & "\"
            Arco.Utils.Logging.Log(vbTab & "Tempdir = " & tempdir)

            Arco.Utils.Logging.Log(vbTab & "Create Borderel and add to zip", LogFileNaam)
            Dim lsxmlname As String
            Dim zip As ZipFile = New ZipFile
            lsxmlname = tempdir & "borderel.xml"

            Dim settings As New XmlWriterSettings()
            settings.Indent = True

            Const schemaLocation As String = "http://MFT-01-00.abb.vlaanderen.be/Borderel Borderel.xsd"
            Const ns1 As String = "http://MFT-01-00.abb.vlaanderen.be/Borderel"

            Dim XmlWrt As XmlWriter = XmlWriter.Create(lsxmlname, settings)
            Dim lipos As Integer = 0

            With XmlWrt
                .WriteStartDocument()
                .WriteStartElement("n1", "Borderel", ns1)
                Dim prefix As String = .LookupPrefix("n1")

                .WriteAttributeString("xsi", "schemaLocation", XmlSchema.InstanceNamespace, schemaLocation)
                .WriteStartElement("Bestanden", ns1)
                Dim loobject As Arco.Doma.Library.baseObjects.DM_OBJECT = WFCurrentCase.TargetObject

                For Each loFileInfo As Arco.Doma.Library.FileList.FileInfo In loobject.Files
                    Dim liFileServerID As Integer = loFileInfo.FILESERVER_ID
                    Dim loFsFile As Arco.Doma.Library.FileServers.DM_FileServer = Arco.Doma.Library.FileServers.DM_FileServer.GetFileServerByID(liFileServerID)
                    Dim loFile As Arco.Doma.Library.File = Arco.Doma.Library.File.GetFile(loFileInfo.FILE_ID)
                    .WriteStartElement("Bestand")
                    .WriteStartElement("Bestandsnaam")
                    lipos = loFile.FILE_PATH.LastIndexOf("\")
                    .WriteString(loFile.FILE_PATH.Substring(lipos + 1))
                    .WriteEndElement() 'Bestandsnaam
                    .WriteEndElement() 'Bestand
                    Arco.Utils.Logging.Debug("File " & loFsFile.RealBasePath & loFile.FILE_PATH & " is added to zip")
                    zip.AddFile(loFsFile.RealBasePath & loFile.FILE_PATH, "")
                Next

                .WriteEndElement() 'Bestanden
                .WriteStartElement("RouteringsMetadata", ns1)
                .WriteElementString("Entiteit", "ABB")
                .WriteElementString("Toepassing", "BBC DR")
                .WriteStartElement("ParameterSet")
                .WriteStartElement("ParameterParameterWaarde")
                .WriteElementString("Parameter", "SLEUTEL")
                .WriteElementString("ParameterWaarde", WFCurrentCase.GetProperty("ondernemingsnummer"))
                .WriteEndElement() 'ParameterParameterWaarde
                .WriteStartElement("ParameterParameterWaarde")
                .WriteElementString("Parameter", "FLOW")
                .WriteElementString("ParameterWaarde", "AANLEVERING GEDAAN")
                .WriteEndElement() 'ParameterParameterWaarde
                .WriteEndElement() 'ParameterSet
                .WriteEndElement() 'RouteringsMetadata
                .WriteEndElement() 'Borderel
                .WriteEndDocument()
                .Close()
            End With

            zip.AddFile(lsxmlname, "")
            zip.Save(mftdir & "AanVlaanderen\" & "BBCDR_AANLEVERINGGEDAAN_" & datetimeexport & ".zip")
            WFCurrentCase.SetProperty("ReferentieZIP", "BBCDR_AANLEVERINGGEDAAN_" & datetimeexport & ".zip")
            Arco.Utils.Logging.Log(vbTab & "verwijderen van " & lsxmlname, LogFileNaam)
            IO.File.Delete(lsxmlname)
            Arco.Utils.Logging.Log(vbTab & "verwijderen van " & tempdir, LogFileNaam)
            IO.Directory.Delete(tempdir)
            zip.Dispose()

        Catch ex As Exception
            bEndedSuccessFully = False
            Arco.Utils.Logging.LogError("Error in " & HuidigeKlasse & "-" & HuidigeMethod & ":" & ex.Message)
            Arco.Utils.Logging.Log(vbTab & "Er is een fout opgetreden in " & HuidigeKlasse & "-" & HuidigeMethod & ". Check de errorlogs.", LogFileNaam)
        End Try
        Arco.Utils.Logging.Log(vbTab & HuidigeMethod & " beëindigd.", LogFileNaam)

    End Sub

    ' DBe 2014/10/13: Niet meer uitgevoerde code. Code wordt in commentaar geplaatst voor de duidelijkheid.
    '                 ************************************************************************************** 
    '' ''Sub getdatafromfile(ByVal WFCurrentCase As cCase, ByVal vsfilename As String)

    '' ''    Dim globalXML As New XmlDocument()
    '' ''    'Dim namespaces As XmlNamespaceManager = New XmlNamespaceManager(globalXML.NameTable)

    '' ''    'namespaces.AddNamespace("aan", "http://BBC-DR-01-00.abb.vlaanderen.be/Aanlevering")
    '' ''    globalXML.Load(vsfilename)

    '' ''    'create namespacemanager with the namespace used in the xml document
    '' ''    Dim namespaces As XmlNamespaceManager
    '' ''    namespaces = createNsMgrForDocument(globalXML)
    '' ''    Dim prefix As String
    '' ''    prefix = namespaces.LookupPrefix("http://BBC-DR-01-00.abb.vlaanderen.be/Aanlevering")


    '' ''    Dim Aanlnode As XmlNode = globalXML.SelectSingleNode("//" & prefix & ":Aanlevering", namespaces)
    '' ''    Arco.Utils.Logging.Log("get data " & vsfilename)
    '' ''    If Aanlnode IsNot Nothing Then
    '' ''        Arco.Utils.Logging.Log("aanlevering found")
    '' ''        If Aanlnode.Item(prefix & ":Ondernemingsnummer") IsNot Nothing Then
    '' ''            Arco.Utils.Logging.Log("ondern nr gevonden" & Aanlnode.Item(prefix & ":Ondernemingsnummer").InnerText)
    '' ''        Else
    '' ''            Arco.Utils.Logging.Log("ondern niet gevonden")
    '' ''        End If
    '' ''        If WFCurrentCase.GetProperty("Ondernemingsnummer") <> Aanlnode.Item(prefix & ":Ondernemingsnummer").InnerText Then
    '' ''            WFCurrentCase.RejectComment = "Ondernemingsnummer in xml bestand stemt niet overeen met het eigen ondernemingsnummer."
    '' ''            WFCurrentCase.RejectUser = "BBCDR"
    '' ''            WFCurrentCase.SetProperty("Rejected", True)
    '' ''            WFCurrentCase.Save()
    '' ''            Exit Sub
    '' ''        Else
    '' ''            WFCurrentCase.RejectComment = ""
    '' ''            WFCurrentCase.RejectUser = ""
    '' ''            WFCurrentCase.SetProperty("Rejected", False)
    '' ''            WFCurrentCase.Save()
    '' ''        End If
    '' ''        WFCurrentCase.SetProperty("Rapport", Aanlnode.Item(prefix & ":RapportCode").InnerText)
    '' ''        WFCurrentCase.SetProperty("Status rekening", Aanlnode.Item(prefix & ":StatusCode").InnerText)
    '' ''        WFCurrentCase.SetProperty("Boekjaar", Aanlnode.Item(prefix & ":Boekjaar").InnerText)
    '' ''        WFCurrentCase.Case_Name = Aanlnode.Item(prefix & ":RapportCode").InnerText & "_" & Aanlnode.Item(prefix & ":StatusCode").InnerText & "_" & Aanlnode.Item(prefix & ":Boekjaar").InnerText
    '' ''        WFCurrentCase.TargetObject.Name = Aanlnode.Item(prefix & ":RapportCode").InnerText & "_" & Aanlnode.Item(prefix & ":StatusCode").InnerText & "_" & Aanlnode.Item(prefix & ":Boekjaar").InnerText
    '' ''    End If
    '' ''    Arco.Utils.Logging.Log("after endif")
    '' ''    globalXML = Nothing
    '' ''End Sub

    ' DBe 2014/10/13: Niet meer uitgevoerde code. Code wordt in commentaar geplaatst voor de duidelijkheid.
    '                 ************************************************************************************** 
    '' ''Private Function GetVariable(ByVal vsKey As String) As Object
    '' ''    If Arco.Doma.Library.Security.BusinessIdentity.CurrentIdentity.Properties.ContainsKey(vsKey) Then
    '' ''        Return Arco.Doma.Library.Security.BusinessIdentity.CurrentIdentity.Properties.Item(vsKey)
    '' ''    Else
    '' ''        Return Nothing
    '' ''    End If
    '' ''End Function



    ' DBe 2014/10/13: Niet meer uitgevoerde code. Code wordt in commentaar geplaatst voor de duidelijkheid.
    '                 ************************************************************************************** 
    '' ''Create an XmlNamespaceManager based on a source XmlDocument's name table, 
    '' '' and prepopulates its namespaces with any 'xmlns:' attributes of the root node.
    '' ''<param name="vsourceDocument">The source XML document to create the XmlNamespaceManager for.</param>
    '' ''<returns>The created XmlNamespaceManager.</returns>

    '' ''Private Function createNsMgrForDocument(ByVal vsourceDocument As XmlDocument) As XmlNamespaceManager

    '' ''    Dim nsmgr As XmlNamespaceManager = New XmlNamespaceManager(vsourceDocument.NameTable)
    '' ''    For Each attr As XmlAttribute In vsourceDocument.SelectSingleNode("/*").Attributes

    '' ''        If (attr.Prefix = "xmlns") Then
    '' ''            nsmgr.AddNamespace(attr.LocalName, attr.Value)
    '' ''        End If
    '' ''    Next

    '' ''    Return nsmgr
    '' ''End Function
#End Region

End Class
