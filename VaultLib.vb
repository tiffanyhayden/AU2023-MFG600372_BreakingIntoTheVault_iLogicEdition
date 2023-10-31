
Imports Connectivity.InventorAddin.EdmAddin
Imports ACW = Autodesk.Connectivity.WebServices
Imports VDF = Autodesk.DataManagement.Client.Framework
Imports Autodesk.DataManagement.Client.Framework.Vault.Currency.Connections
Imports Autodesk.DataManagement.Client.Framework.Vault.Currency.Properties
Imports Autodesk.DataManagement.Client.Framework.Vault.Currency.Entities
Imports Autodesk.DataManagement.Client.Framework.Vault.Services
Imports System.IO

Public Class VaultLib

    Dim oEDMS As EdmSecurity = EdmSecurity.Instance
    Dim oConnection As Connection = oEDMS.VaultConnection

#Region "_01 Vault_IsConnected"
    Public Function Vault_IsConnected() As Boolean

        If oConnection IsNot Nothing Then Vault_IsConnected = True : Exit Function
        Vault_IsConnected = False

        Try
            '*******************************************************************************************
            ' ESTABLISH EDM CONNECTION
            '*******************************************************************************************
            If oEDMS Is Nothing Then oEDMS = EdmSecurity.Instance

            '*******************************************************************************************
            ' CONNECT TO VAULT ADD IN IF NEEDED
            '*******************************************************************************************
            If oConnection Is Nothing Then oConnection = oEDMS.VaultConnection
            If oConnection IsNot Nothing Then Vault_IsConnected = True
        Catch EX As Exception
            MsgBox("Connection Failed. Check Inventor Vault Add-in", MsgBoxStyle.OkOnly, "Error")
            Exit Function
        End Try
    End Function

#End Region

#Region "_02 File_GetByFileName"

    Function File_GetByFileName(strFileName As String, Optional blnDoCheckOut As Boolean = False, Optional strSearchIn As String = "$/") As String
        '*******************************************************************************************
        ' GIVEN A FILENAME THE FILE IS DOWNLOADED LOCALLY USING THE VAULT ADDIN CONNECTION.
        '*******************************************************************************************
        '*******************************************************************************************
        ' CHECK MAJOR REQUIREMENTS FOR THIS FUNCTION AND EXITS EARLY IF NEEDED. 
        '*******************************************************************************************
        File_GetByFileName = ""
        If strFileName = "" Then
            MsgBox("strFileName was empty string, value must be defined.", MsgBoxStyle.OkOnly, "Error")
            Exit Function
        End If

        Dim oInvApp As Inventor.Application = System.Runtime.InteropServices.Marshal.GetActiveObject("Inventor.Application")
        Dim oFile As ACW.File
        Dim oFileIteration As VDF.Vault.Currency.Entities.FileIteration
        Dim oServices As VDF.Vault.Services.Connection.IWorkingFoldersManager


        '*******************************************************************************************
        ' CHECK CONNECTION. EXIT EARLY IF FAILURE
        '*******************************************************************************************
        If Vault_IsConnected() = False Then Exit Function

        '*******************************************************************************************
        ' USE CONNECTION AND FILE NAME TO ESTABLISH THE FILE
        '*******************************************************************************************
        Try
            oFile = File_FindByFilename(strFileName,, strSearchIn)
        Catch EX As Exception
            MsgBox("File was not found, please insure the file name is correct", MsgBoxStyle.OkOnly, "Error")
            File_GetByFileName = ""
            Exit Function
        End Try


        '*******************************************************************************************
        ' USE THE FILE AND CONNECTION TO CREATE A FILE ITERATION AND SET FULL FILE PATH
        '*******************************************************************************************
        Try
            oFileIteration = New VDF.Vault.Currency.Entities.FileIteration(oConnection, oFile)
            oServices = oConnection.WorkingFoldersManager
            File_GetByFileName = oServices.GetPathOfFileInWorkingFolder(oFileIteration).FullPath.ToString
        Catch EX As Exception
            MsgBox("File connection failed.", MsgBoxStyle.OkOnly, "Error")
            File_GetByFileName = ""
            Exit Function
        End Try

        '*******************************************************************************************
        ' USE THE FILE AND CONNECTION TO CREATE A FILE ITERATION AND SET FULL FILE PATH
        '*******************************************************************************************
        Try
            File_Acquire(oFile, blnDoCheckOut)
        Catch ex As Exception
            MsgBox("Possible Vault connection issue, file exists in Vault but were unable to aquire file.", MsgBoxStyle.OkOnly, "Error")
            File_GetByFileName = ""
        End Try



    End Function
    Function File_FindByFilename(strFilename As String, Optional intSrchOper As Integer = 3, Optional strSearchIn As String = "$/") As ACW.File
        File_FindByFilename = Nothing

        '*******************************************************************************************
        ' ESTABLISH VAULT CONNECTION USING INVENTOR VAULT ADD-IN
        '*******************************************************************************************
        If Vault_IsConnected() = False Then Exit Function

        Dim oDocService As ACW.DocumentService
        Dim oFile As ACW.File = Nothing
        Dim oPropDefs As ACW.PropDef() = oConnection.WebServiceManager.PropertyService.GetPropertyDefinitionsByEntityClassId("FILE")
        Dim oPropDef As ACW.PropDef = oPropDefs.[Single](Function(n) n.SysName = "ClientFileName")
        Dim oSearch As New ACW.SrchCond()
        Dim oFolder As ACW.Folder = Nothing
        Dim oFolders As ACW.Folder() = Nothing
        Dim strBookmark As String = String.Empty
        Dim oStatus As ACW.SrchStatus = Nothing
        Dim oResults As ACW.File()
        Dim lgFolderIds As Long() = Nothing

        '*******************************************************************************************
        ' ESTABLISH DOCUMENT SERVICE CONNECTION USING ABOVE CONNECTION
        '*******************************************************************************************
        Try
            oDocService = oConnection.WebServiceManager.DocumentService
        Catch EX As Exception
            MsgBox("Document Service Failed to establish connection.", MsgBoxStyle.OkOnly, "Error")
            Exit Function
        End Try

        '*******************************************************************************************
        ' DEFINE SEARCH CONDITION OPTIONS
        '*******************************************************************************************
        Try
            oSearch.PropDefId = oPropDef.Id
            oSearch.PropTyp = ACW.PropertySearchType.SingleProperty
            oSearch.SrchOper = intSrchOper
            oSearch.SrchRule = ACW.SearchRuleType.Must
            oSearch.SrchTxt = strFilename
        Catch EX As Exception
            MsgBox("Search criteria failed.", MsgBoxStyle.OkOnly, "Error")
            Exit Function
        End Try

        '*******************************************************************************************
        ' FIND FOLDERS BY PATH
        '*******************************************************************************************
        Try
            oFolders = oDocService.FindFoldersByPaths({strSearchIn})
        Catch EX As Exception
            MsgBox("Finding folders failed.", MsgBoxStyle.OkOnly, "Error")
            Exit Function
        End Try

        '*******************************************************************************************
        ' LOOP THROUGH FOLDERS AND ADD ID TO LONG ARRAY
        '*******************************************************************************************
        For Each oFolder In oFolders
            If oFolder.Id <> -1 Then
                lgFolderIds = {oFolder.Id}
            End If
        Next

        '*******************************************************************************************
        ' SET LONG ARRAY TO NOTHING IF NO IDs ARE FOUND
        '*******************************************************************************************
        If lgFolderIds.Length = 0 Then lgFolderIds = {Nothing}

        '*******************************************************************************************
        ' CREATE SEARCH
        '*******************************************************************************************
        Try
            oResults = oConnection.WebServiceManager.DocumentService.FindFilesBySearchConditions(New ACW.SrchCond() {oSearch}, Nothing, lgFolderIds, True, True, strBookmark, oStatus)
        Catch EX As Exception
            MsgBox("Search Failed.", MsgBoxStyle.OkOnly, "Error")
            Exit Function
        End Try

        '*******************************************************************************************
        ' CREATE SEARCH
        '*******************************************************************************************
        Try
            File_FindByFilename = oResults(0)
        Catch EX As Exception
            MsgBox("No file returned, check file name accuracy.", MsgBoxStyle.OkOnly, "Error")
            Exit Function
        End Try

    End Function
    Sub File_Acquire(oFile As ACW.File, Optional blnDoCheckOut As Boolean = False)

        Dim oParent As New IntPtr
        Dim oSettings As VDF.Vault.Forms.Settings.InteractiveAcquireFileSettings
        Dim oFileIteration As VDF.Vault.Currency.Entities.FileIteration

        '*******************************************************************************************
        ' ESTABLISH VAULT CONNECTION USING INVENTOR VAULT ADD-IN
        '*******************************************************************************************
        If Vault_IsConnected() = False Then Exit Sub

        '*******************************************************************************************
        ' CREATE FILE ITERATION
        '*******************************************************************************************
        Try
            oFileIteration = New VDF.Vault.Currency.Entities.FileIteration(oConnection, oFile)
        Catch ex As Exception
            MsgBox("File iteration not created", MsgBoxStyle.OkOnly, "Error")
            Exit Sub
        End Try

        '*******************************************************************************************
        ' CREATE FILE ITERATION
        '*******************************************************************************************
        Try
            oSettings = New VDF.Vault.Forms.Settings.InteractiveAcquireFileSettings(oConnection, oParent, "Download files")
        Catch ex As Exception
            MsgBox("Settings failed.", MsgBoxStyle.OkOnly, "Error")
            Exit Sub
        End Try

        '*******************************************************************************************
        ' DEFINE SETTING OPTIONS
        '*******************************************************************************************
        oSettings.OptionsResolution.OverwriteOption = VDF.Vault.Settings.AcquireFilesSettings.AcquireFileResolutionOptions.OverwriteOptions.ForceOverwriteAll
        oSettings.OptionsResolution.SyncWithRemoteSiteSetting = VDF.Vault.Settings.AcquireFilesSettings.SyncWithRemoteSite.Always
        oSettings.OptionsRelationshipGathering.FileRelationshipSettings.IncludeAttachments = False
        oSettings.OptionsRelationshipGathering.FileRelationshipSettings.IncludeChildren = False
        oSettings.OptionsRelationshipGathering.FileRelationshipSettings.IncludeHiddenEntities = False
        oSettings.OptionsRelationshipGathering.FileRelationshipSettings.IncludeLibraryContents = False
        oSettings.OptionsRelationshipGathering.FileRelationshipSettings.IncludeParents = False
        oSettings.OptionsRelationshipGathering.FileRelationshipSettings.IncludeRelatedDocumentation = False
        oSettings.OptionsRelationshipGathering.FileRelationshipSettings.RecurseChildren = False
        oSettings.OptionsRelationshipGathering.FileRelationshipSettings.RecurseParents = False
        oSettings.OptionsRelationshipGathering.FileRelationshipSettings.ReleaseBiased = False

        '*******************************************************************************************
        ' SET DOWNLOAD OR CHECKOUT FILE SETTING. ADD TO ACQUIRE FILE
        '*******************************************************************************************
        If blnDoCheckOut = True Then
            oSettings.DefaultAcquisitionOption = VDF.Vault.Settings.AcquireFilesSettings.AcquisitionOption.Checkout
        Else
            oSettings.DefaultAcquisitionOption = VDF.Vault.Settings.AcquireFilesSettings.AcquisitionOption.Download
        End If

        '*******************************************************************************************
        ' ACQUIRE FILE
        '*******************************************************************************************
        Try
            oSettings.AddFileToAcquire(oFileIteration, oSettings.DefaultAcquisitionOption)
            oConnection.FileManager.AcquireFiles(oSettings)
        Catch ex As Exception
            MsgBox("Acquire file failed.", MsgBoxStyle.OkOnly, "Error")
            Exit Sub
        End Try

    End Sub


#End Region

#Region "_03 Folder_GetByPath"

    Function Folder_GetByPath(strFolderLocalPath As String, Optional blnDeleteDirectory As Boolean = False,
                              Optional strLocalWorkingFolder As String = "C:/_VAULT2019") As String
        '*******************************************************************************************
        ' CHECK MAJOR REQUIREMENTS FOR THIS FUNCTION AND EXITS EARLY IF NEEDED. 
        '*******************************************************************************************
        Folder_GetByPath = ""
        If strFolderLocalPath = "" Then
            MsgBox("strFullFolderPath must not be an empty string, try another string value.", MsgBoxStyle.OkOnly, "Error")
            Exit Function
        End If

        Dim oProcess As New ProcessStartInfo
        Dim oInvApp As Inventor.Application = System.Runtime.InteropServices.Marshal.GetActiveObject("Inventor.Application")
        Dim strFullPathSplit As String()
        Dim strFolderName As String
        Dim oFolder As ACW.Folder
        Dim oFolders As ACW.Folder()
        Dim oFolderFound As ACW.Folder = Nothing
        Dim oFolderiEntity As VDF.Vault.Currency.Entities.IEntity = Nothing
        Dim strVaultPath As String

        '*******************************************************************************************
        ' ESTABLISH VAULT CONNECTION USING INVENTOR VAULT ADD-IN
        '*******************************************************************************************
        If Vault_IsConnected() = False Then Exit Function

        '**********************************************************
        'CHANGES FULLFOLDERPATH IF FOR SOME REASON SOMEONE INCLUDES
        'THE LAST \ IN THE FILENAME BY MISTAKE
        '**********************************************************
        If strFolderLocalPath.EndsWith("\") Then strFolderLocalPath = Left(strFolderLocalPath, Len(strFolderLocalPath) - 1)
        strVaultPath = strFolderLocalPath

        '**********************************************************
        ' CLEARS THE DIRECTORY IF REQUESTED
        '**********************************************************

        If blnDeleteDirectory = True Then
            Try
                LocalFolder_ClearContents(strFolderLocalPath)
            Catch
                MsgBox("Clearing the directory was not possible.", MsgBoxStyle.OkOnly, "Error")
            End Try
        End If

        '***************************************************************
        ' FIND THE FOLDER NAME
        '***************************************************************
        strFullPathSplit = Split(strVaultPath, "\")
        strFolderName = strFullPathSplit.Last

        '***************************************************************
        ' CLEAN UP WORKING FOLDER PATH TO MATCH VAULT PATH SCHEME
        '***************************************************************
        strVaultPath = Replace(strVaultPath, "\", "/")
        strVaultPath = Replace(strVaultPath, strLocalWorkingFolder, "$")
        '***************************************************************
        ' FINDS THE FOLDER BASED OFF OF THE FULL PATH
        '***************************************************************
        oFolders = Folders_FindByName(strFolderName)

        For Each oFolder In oFolders
            If oFolder.FullName = strVaultPath Then
                oFolderFound = oFolder
                Exit For
            End If
        Next

        '***************************************************************
        'CONVERTS THE FOLDER TO A FOLDER ENTITY
        '***************************************************************

        Try
            If oFolderFound IsNot Nothing Then oFolderiEntity = New VDF.Vault.Currency.Entities.Folder(oConnection, oFolderFound)
        Catch
            MsgBox("Folder iEntity could not be created", MsgBoxStyle.OkOnly, "Error")
            Exit Function
        End Try

        '***************************************************************
        'ACQUIRES THE FOLDER AND CONTENTS IF A FOLDER IS FOUND. 
        '***************************************************************
        Try
            If oFolderiEntity IsNot Nothing Then
                Folder_Acquire(strFolderLocalPath, oFolderiEntity)
                Folder_GetByPath = strFolderLocalPath
            End If
        Catch
            MsgBox("Folder could not be acquired.", MsgBoxStyle.OkOnly, "Error")
            Exit Function
        End Try

    End Function

    Function Folders_FindByName(strFolderName As String, Optional intSrchOper As Integer = 3) As ACW.Folder()
        '*******************************************************************************************
        ' THIS FUNCTION RETURNS A LIST OF FOLDERS BASED ON FOLDER NAME
        ' IF MORE THAN ONE IS FOUND, MORE THAN ONE IS RETURNED
        '*******************************************************************************************


        '*******************************************************************************************
        ' CHECK MAJOR REQUIREMENTS FOR THIS FUNCTION AND EXITS EARLY IF NEEDED. 
        '*******************************************************************************************
        Folders_FindByName = Nothing
        If strFolderName = "" Then
            MsgBox("strFolderName was empty string, must be defined. Please define one.", MsgBoxStyle.OkOnly, "Error")
            Exit Function
        End If

        Dim oDocService As ACW.DocumentService
        Dim oFolders As ACW.Folder() = Nothing
        Dim oPropDefs As ACW.PropDef()
        Dim oPropDef As ACW.PropDef
        Dim oSearch As New ACW.SrchCond()
        Dim strBookMark As String
        Dim oStatus As ACW.SrchStatus = Nothing

        '*******************************************************************************************
        ' ESTABLISH VAULT CONNECTION USING INVENTOR VAULT ADD-IN
        '*******************************************************************************************
        If Vault_IsConnected() = False Then Exit Function


        oDocService = oConnection.WebServiceManager.DocumentService
        oPropDefs = oConnection.WebServiceManager.PropertyService.GetPropertyDefinitionsByEntityClassId("FLDR")
        oPropDef = oPropDefs.[Single](Function(n) n.DispName = "Name")
        strBookMark = String.Empty


        oSearch.PropDefId = oPropDef.Id
        oSearch.PropTyp = ACW.PropertySearchType.SingleProperty
        oSearch.SrchOper = intSrchOper
        oSearch.SrchRule = ACW.SearchRuleType.Must
        oSearch.SrchTxt = strFolderName

        Try
            Folders_FindByName = oConnection.WebServiceManager.DocumentService.FindFoldersBySearchConditions(New ACW.SrchCond() {oSearch}, Nothing, {}, True, strBookMark, oStatus)
        Catch ex As Exception
            MsgBox("Folders were not found", MsgBoxStyle.OkOnly, "Error")
            Exit Function
        End Try



    End Function

    Sub Folder_Acquire(strLocalPath As String, Optional oFolderiEntity As VDF.Vault.Currency.Entities.IEntity = Nothing, Optional oFolder As ACW.Folder = Nothing)

        Dim oParent As New IntPtr
        Dim oSettings As VDF.Vault.Forms.Settings.InteractiveAcquireFileSettings

        '*******************************************************************************************
        ' ESTABLISH VAULT CONNECTION USING INVENTOR VAULT ADD-IN
        '*******************************************************************************************
        If Vault_IsConnected() = False Then Exit Sub

        '*******************************************************************************************
        ' CREATE FOLDER ENTITY
        '*******************************************************************************************

        Try
            If oFolderiEntity Is Nothing And oFolder IsNot Nothing Then oFolderiEntity = New VDF.Vault.Currency.Entities.Folder(oConnection, oFolder)
        Catch EX As Exception
            MsgBox("Folder iEntity could not be created. ", MsgBoxStyle.OkOnly, "Error")
            Exit Sub
        End Try

        '*******************************************************************************************
        ' DEFINE SETTINGS
        '*******************************************************************************************
        oSettings = New VDF.Vault.Forms.Settings.InteractiveAcquireFileSettings(oConnection, oParent, "Download files")
        oSettings.OptionsResolution.OverwriteOption = VDF.Vault.Settings.AcquireFilesSettings.AcquireFileResolutionOptions.OverwriteOptions.ForceOverwriteAll
        oSettings.DefaultAcquisitionOption = VDF.Vault.Settings.AcquireFilesSettings.AcquisitionOption.Download
        oSettings.OptionsResolution.SyncWithRemoteSiteSetting = VDF.Vault.Settings.AcquireFilesSettings.SyncWithRemoteSite.Always

        oSettings.OptionsRelationshipGathering.FileRelationshipSettings.IncludeAttachments = False
        oSettings.OptionsRelationshipGathering.FileRelationshipSettings.IncludeChildren = False
        oSettings.OptionsRelationshipGathering.FileRelationshipSettings.IncludeHiddenEntities = False
        oSettings.OptionsRelationshipGathering.FileRelationshipSettings.IncludeLibraryContents = False
        oSettings.OptionsRelationshipGathering.FileRelationshipSettings.IncludeParents = False
        oSettings.OptionsRelationshipGathering.FileRelationshipSettings.IncludeRelatedDocumentation = False
        oSettings.OptionsRelationshipGathering.FileRelationshipSettings.RecurseChildren = False
        oSettings.OptionsRelationshipGathering.FileRelationshipSettings.RecurseParents = False
        oSettings.OptionsRelationshipGathering.FileRelationshipSettings.ReleaseBiased = False

        oSettings.AddEntityToAcquire(oFolderiEntity)
        oSettings.LocalPath = New VDF.Currency.FolderPathAbsolute(strLocalPath)

        '*******************************************************************************************
        ' ACQUIRE FILE
        '*******************************************************************************************
        Try
            oConnection.FileManager.AcquireFiles(oSettings)
        Catch EX As Exception
            MsgBox("Acquiring file failed.", MsgBoxStyle.OkOnly, "Error")
            Exit Sub
        End Try




    End Sub

#End Region

#Region "_04 Folder_ExportByExtension"

    Public Function Folder_ExportByExtension(strFolderFullPath As String, strFileExtension As String, Optional strTargetPath As String = "", Optional strLocalWorkingFolder As String = "C:\_VAULT2019") As String



        '*******************************************************************************************
        ' CHECK MAJOR REQUIREMENTS FOR THIS FUNCTION AND EXITS EARLY IF NEEDED. 
        '*******************************************************************************************
        'strFileExtension = Extension examples include: "ALL", "IPT", "IAM", "DWG", or any other extension.
        Folder_ExportByExtension = Nothing
        If strFolderFullPath = "" And strFileExtension <> "" Then
            MsgBox("strFolderFullPath AND strFileExtension must have a value to continue", MsgBoxStyle.OkOnly, "Error")
            Exit Function
        End If

        Dim oTextFile As StreamWriter
        Dim strFileName As String = ""
        Dim oFile As ACW.File
        Dim oFiles As ACW.File()

        '*******************************************************************************************
        ' ESTABLISH VAULT CONNECTION USING VAULT ADD-IN WITHIN INVENTOR
        '*******************************************************************************************
        If Vault_IsConnected() = False Then Exit Function

        '*******************************************************************************************
        ' ADJUST VAULT FILE PATH TO MEET STANDARD PATH NAMING SCHEME
        '*******************************************************************************************
        If strFolderFullPath.StartsWith("C:\") Then strFolderFullPath = strFolderFullPath.Replace(strLocalWorkingFolder, "$") ' CHANGE WORKING FOLDER NAME. YOURS WILL BE DIFFERENT
        If strFolderFullPath.EndsWith("\") Then strFolderFullPath = strFolderFullPath.TrimEnd(CChar("\"))
        If strFolderFullPath.EndsWith("/") Then strFolderFullPath = strFolderFullPath.TrimEnd(CChar("/"))
        If InStr(strFolderFullPath, "\") Then strFolderFullPath = strFolderFullPath.Replace("\", "/")

        '*******************************************************************************************
        ' GATHER FILES IN A DETERMINED FOLDER
        '*******************************************************************************************
        Try
            oFiles = Files_FindByFolderPath(strFolderFullPath, 3)
        Catch ex As Exception
            MsgBox("Gathering files was not possible", MsgBoxStyle.OkOnly, "Error")
            Exit Function
        End Try


        '*******************************************************************************************
        ' GATHER ALL THE FILES THAT MATCH THE TARGET EXTENSION
        '*******************************************************************************************
        If oFiles IsNot Nothing Then
            For Each oFile In oFiles
                If oFile.Hidden = False Then
                    If strFileExtension = "ALL" Then
                        strFileName &= oFile.Name & vbCrLf
                    Else
                        If oFile.Name.ToUpper.EndsWith(strFileExtension.ToUpper) = True Then
                            strFileName &= oFile.Name & vbCrLf
                        End If
                    End If
                End If
            Next
        End If

        '*******************************************************************************************
        ' WRITE TO THE TEXT FILE IF TARGET PATH IS NOT EMPTY
        '*******************************************************************************************
        If strTargetPath <> "" Then
            If System.IO.File.Exists(strTargetPath) = True Then
                System.IO.File.WriteAllText(strTargetPath, strFileName)
            Else
                System.IO.File.Create(strTargetPath).Dispose()
                oTextFile = New System.IO.StreamWriter(strTargetPath)
                oTextFile.WriteLine(strFileName)
                oTextFile.Close()
            End If
        End If

        Folder_ExportByExtension = strFileName


    End Function

    Public Function Files_FindByFolderPath(strVaultFolderPath As String, Optional intSrchOper As Integer = 1) As ACW.File()



        '*******************************************************************************************
        ' CHECK MAJOR REQUIREMENTS FOR THIS FUNCTION AND EXITS EARLY IF NEEDED. 
        '*******************************************************************************************
        Files_FindByFolderPath = Nothing
        If strVaultFolderPath = "" Then
            MsgBox("strVaultFolderPath must have a value to continue", MsgBoxStyle.OkOnly, "Error")
            Exit Function
        End If

        Dim oFolder As ACW.Folder = Nothing
        Dim oPropDefs As ACW.PropDef()
        Dim oPropDef As ACW.PropDef
        Dim oSearch As New ACW.SrchCond()
        Dim strBookmark As String
        Dim oStatus As ACW.SrchStatus = Nothing
        Dim oFiles As ACW.File()

        '*******************************************************************************************
        ' ESTABLISH VAULT CONNECTION USING VAULT ADD-IN WITHIN INVENTOR
        '*******************************************************************************************
        If Vault_IsConnected() = False Then Exit Function

        '*******************************************************************************************
        ' EXTERNAL RULES FOLDER
        '*******************************************************************************************
        Try
            If oConnection IsNot Nothing Then oFolder = Folder_FindByPath(strVaultFolderPath, intSrchOper)
        Catch ex As Exception
            MsgBox("Folder could not be found.", MsgBoxStyle.OkOnly, "Error")
            Exit Function
        End Try

        '*******************************************************************************************
        ' START SEARCH CONDITIONS/PROPS
        '*******************************************************************************************
        oPropDefs = oConnection.WebServiceManager.PropertyService.GetPropertyDefinitionsByEntityClassId("FILE")
        oPropDef = oPropDefs.[Single](Function(n) n.DispName = "Folder Path")

        oSearch.PropDefId = oPropDef.Id
        oSearch.PropTyp = ACW.PropertySearchType.SingleProperty
        oSearch.SrchOper = intSrchOper
        oSearch.SrchRule = ACW.SearchRuleType.Must
        oSearch.SrchTxt = strVaultFolderPath
        strBookmark = String.Empty

        '*******************************************************************************************
        ' EXTERNAL RULES FOLDER
        '*******************************************************************************************
        Try
            oFiles = oConnection.WebServiceManager.DocumentService.FindFilesBySearchConditions(New ACW.SrchCond() {oSearch}, Nothing, {oFolder.Id}, True, True, strBookmark, oStatus)
        Catch ex As Exception
            MsgBox("Files were not found.", MsgBoxStyle.OkOnly, "Error")
            Exit Function
        End Try


        If oFiles IsNot Nothing Then Files_FindByFolderPath = oFiles

    End Function

    Function Folder_FindByPath(strVaultFolderPath As String, Optional intSrchOper As Integer = 1) As ACW.Folder
        '*******************************************************************************************
        ' GIVEN A CONNECTION AND FOLDER NAME FIND A FOLDER IN VAULT
        '*******************************************************************************************
        Folder_FindByPath = Nothing
        Dim oDocService As ACW.DocumentService = Nothing
        Dim oPropDefs As ACW.PropDef()
        Dim oPropDef As ACW.PropDef
        Dim oSearch As New ACW.SrchCond()
        Dim oFolders As ACW.Folder()
        Dim oFolder As ACW.Folder = Nothing

        '*******************************************************************************************
        ' ESTABLISH VAULT CONNECTION USING INVENTOR VAULT ADD-IN
        '*******************************************************************************************
        If Vault_IsConnected() = False Then Exit Function

        '*******************************************************************************************
        ' ESTABLISH DOCUMENT SERVICE CONNECTION USING ABOVE CONNECTION
        '*******************************************************************************************
        Try
            oDocService = oConnection.WebServiceManager.DocumentService
        Catch EX As Exception
            MsgBox("Document Service Failed to establish connection.", MsgBoxStyle.OkOnly, "Error")
            Exit Function
        End Try

        '*******************************************************************************************
        ' CREATE PROP DEF(S) USING CONNECTION.
        '*******************************************************************************************
        Try
            oPropDefs = oConnection.WebServiceManager.PropertyService.GetPropertyDefinitionsByEntityClassId("FLDR")
            oPropDef = oPropDefs.[Single](Function(n) n.DispName = "Folder Path")
        Catch EX As Exception
            MsgBox("Prop Defs failed.", MsgBoxStyle.OkOnly, "Error")
            Exit Function
        End Try

        '*******************************************************************************************
        ' DEFINE SEARCH CONDITION OPTIONS
        '*******************************************************************************************
        Try
            oSearch.PropDefId = oPropDef.Id
            oSearch.PropTyp = ACW.PropertySearchType.SingleProperty
            oSearch.SrchOper = intSrchOper
            oSearch.SrchRule = ACW.SearchRuleType.Must
            oSearch.SrchTxt = strVaultFolderPath
        Catch EX As Exception
            MsgBox("Search criteria failed.", MsgBoxStyle.OkOnly, "Error")
            Exit Function
        End Try

        '*******************************************************************************************
        ' CREATE SEARCH
        '*******************************************************************************************
        Try
            oFolders = oDocService.FindFoldersBySearchConditions(New ACW.SrchCond() {oSearch}, Nothing, Nothing, True, True, Nothing)
        Catch EX As Exception
            MsgBox("Finding folders failed.", MsgBoxStyle.OkOnly, "Error")
            Exit Function
        End Try

        '*******************************************************************************************
        ' FIND FIRST FOLDER FOUND
        '*******************************************************************************************
        Try
            Folder_FindByPath = oFolders(0)
        Catch ex As Exception
            MsgBox("Setting folder failed.", MsgBoxStyle.OkOnly, "Error")
            Exit Function
        End Try

    End Function


#End Region

#Region "_05 File_IsCheckedOut"
    Function File_IsCheckedOut(strFileName As String, Optional blnToCurrentUser As Boolean = False) As Boolean
        '*******************************************************************************************
        ' FUNCTION DETERMINES IF A FILE IS CHECKED OUT
        '*******************************************************************************************

        '*******************************************************************************************
        ' CHECK MAJOR REQUIREMENTS FOR THIS FUNCTION AND EXITS EARLY IF NEEDED. 
        '*******************************************************************************************
        File_IsCheckedOut = False
        If strFileName = "" Then
            MsgBox("oFile was Nothing and strFileName was empty string. Please define one.", MsgBoxStyle.OkOnly, "Error")
            Exit Function
        End If

        Dim oInvApp As Inventor.Application = System.Runtime.InteropServices.Marshal.GetActiveObject("Inventor.Application")
        Dim oFileIteration As VDF.Vault.Currency.Entities.FileIteration
        Dim oFile As ACW.File

        '*******************************************************************************************
        ' ESTABLISH VAULT CONNECTION USING VAULT ADD-IN WITHIN INVENTOR
        '*******************************************************************************************
        If Vault_IsConnected() = False Then Exit Function

        '*******************************************************************************************
        ' FIND FILE 
        '*******************************************************************************************
        Try
            oFile = File_FindByFilename(strFileName)
        Catch ex As Exception
            MsgBox("File was not able to be found, make sure strFileName is correct and Vault Add-in is logged in.", MsgBoxStyle.OkOnly, "Error")
            Exit Function
        End Try


        '*******************************************************************************************
        ' CHECK CHECKED OUT STATUS FOR A FILE
        '*******************************************************************************************
        If oFile IsNot Nothing Then
            oFileIteration = New VDF.Vault.Currency.Entities.FileIteration(oConnection, oFile)
            Try
                If blnToCurrentUser = False Then
                    File_IsCheckedOut = oFileIteration.IsCheckedOut
                ElseIf blnToCurrentUser Then
                    File_IsCheckedOut = oFileIteration.IsCheckedOutToCurrentUser
                End If
            Catch ex As Exception
                MsgBox("Checked out state was not available.", MsgBoxStyle.OkOnly, "Error")
                Exit Function
            End Try
        End If



    End Function

#End Region

#Region "_06 File_UndoCheckOut"

    Function File_UndoCheckOut(strFileName As String) As String
        '*******************************************************************************************
        ' THIS FUNCTION DOES A UNDO CHECKOUT ON A FILE AND RETURNS THE FULL FILE PATH AS A STRING
        ' IF SUCCESSFUL, IF FAILS IT WILL RETURN AND EMTPY STRING
        '*******************************************************************************************

        '*******************************************************************************************
        ' CHECK MAJOR REQUIREMENTS FOR THIS FUNCTION AND EXITS EARLY IF NEEDED. 
        '*******************************************************************************************
        File_UndoCheckOut = ""
        If strFileName = "" Then
            MsgBox("strFileName must have a value to continue", MsgBoxStyle.OkOnly, "Error")
            Exit Function
        End If

        Dim oInvApp As Inventor.Application = System.Runtime.InteropServices.Marshal.GetActiveObject("Inventor.Application")
        Dim strFullFileName As String
        Dim oFileIteration As VDF.Vault.Currency.Entities.FileIteration
        Dim oServices As VDF.Vault.Services.Connection.IWorkingFoldersManager
        Dim oFile As ACW.File

        '*******************************************************************************************
        ' ESTABLISH VAULT CONNECTION USING VAULT ADD-IN WITHIN INVENTOR
        '*******************************************************************************************
        If Vault_IsConnected() = False Then Exit Function


        '*******************************************************************************************
        ' FIND FILE 
        '*******************************************************************************************
        Try
            oFile = File_FindByFilename(strFileName)
        Catch ex As Exception
            MsgBox("File was not able to be found, make sure strFileName is correct and Vault Add-in is logged in.", MsgBoxStyle.OkOnly, "Error")
            Exit Function
        End Try

        '*******************************************************************************************
        ' UNDO CHECKOUT FILE AND RETURN FULL FILE PATH IF SUCCESSFUL
        '*******************************************************************************************
        If oFile IsNot Nothing Then
            oFileIteration = New VDF.Vault.Currency.Entities.FileIteration(oConnection, oFile)
            oServices = oConnection.WorkingFoldersManager
            strFullFileName = oServices.GetPathOfFileInWorkingFolder(oFileIteration).FullPath.ToString

            Try
                If File_IsCheckedOut(strFileName) Then oConnection.FileManager.UndoCheckoutFile(oFileIteration)
                File_UndoCheckOut = strFullFileName
            Catch ex As Exception
                MsgBox("Undocheckout was not available.", MsgBoxStyle.OkOnly, "Error")
            End Try
        End If

    End Function

#End Region

#Region "_07 GetFromFile_LatestVersionCreator"

    Public Function GetFromFile_LatestVersionCreator(strFileName As String) As String
        '*******************************************************************************************
        ' THIS FUNCTION WILL RETURN THE CREATOR OF THE LATEST FILE VERSION 
        '*******************************************************************************************

        '*******************************************************************************************
        ' CHECK MAJOR REQUIREMENTS FOR THIS FUNCTION AND EXITS EARLY IF NEEDED. 
        '*******************************************************************************************
        GetFromFile_LatestVersionCreator = ""
        If strFileName = "" Then
            MsgBox("strFileName must have a value to continue", MsgBoxStyle.OkOnly, "Error")
            Exit Function
        End If

        Dim oFileVersions As ACW.File() = Nothing
        Dim oFileVersion As ACW.File
        Dim oFile As ACW.File
        Dim oDocumentService As ACW.DocumentService = oConnection.WebServiceManager.DocumentService

        '*******************************************************************************************
        ' ESTABLISH VAULT CONNECTION USING VAULT ADD-IN WITHIN INVENTOR
        '*******************************************************************************************
        If Vault_IsConnected() = False Then Exit Function

        '*******************************************************************************************
        ' FIND FILE 
        '*******************************************************************************************
        Try
            oFile = File_FindByFilename(strFileName)
        Catch ex As Exception
            MsgBox("File was not able to be found, make sure strFileName is correct and Vault Add-in is logged in.", MsgBoxStyle.OkOnly, "Error")
            Exit Function
        End Try

        '*******************************************************************************************
        ' GET ALL VERSIONS FOR FILE
        '*******************************************************************************************
        Try
            If oFile IsNot Nothing Then oFileVersions = oDocumentService.GetFilesByMasterId(oFile.MasterId)
        Catch ex As Exception
            MsgBox("File versions were not available.", MsgBoxStyle.OkOnly, "Error")
            Exit Function
        End Try

        '*******************************************************************************************
        ' REVERSE THE ARRAY TO GO FROM EARLIEST VERSION TO LATEST
        '*******************************************************************************************
        Try
            If oFileVersions IsNot Nothing Then Array.Reverse(oFileVersions)
        Catch ex As Exception
            MsgBox("Reversing file versions was not possible.", MsgBoxStyle.OkOnly, "Error")
            Exit Function
        End Try

        '*******************************************************************************************
        ' FIND THE FIRST VESIONS CREATOR THAT IS NOT ADMIN 
        '*******************************************************************************************
        If oFileVersions IsNot Nothing Then
            For Each oFileVersion In oFileVersions
                GetFromFile_LatestVersionCreator = oFileVersion.CreateUserName.ToUpper
                Exit For
            Next
        End If

    End Function

#End Region

#Region "_08 GetFile_LatestVersionByLifecycleState"
    Public Function GetFile_LatestVersionByLifecycleState(strFileName As String, strLifeCycleState As String) As String
        '*******************************************************************************************
        ' THIS FUNCTION "GETS" THE LATEST VERSION OF A FILE THAT IS AT A SPECIFIED LIFECYCLE STATE
        '*******************************************************************************************
        '*******************************************************************************************
        ' CHECK MAJOR REQUIREMENTS FOR THIS FUNCTION AND EXITS EARLY IF NEEDED. 
        '*******************************************************************************************
        strLifeCycleState = strLifeCycleState.ToUpper
        GetFile_LatestVersionByLifecycleState = ""
        If strLifeCycleState = "" Then
            MsgBox("Lifecycle state was not defined, please define and try again.", MsgBoxStyle.OkOnly, "Error")
            Exit Function
        ElseIf strFileName = "" Then
            MsgBox("strFileName must have a value to continue", MsgBoxStyle.OkOnly, "Error")
            Exit Function
        End If

        Dim oFile As ACW.File
        Dim oFileVersions As ACW.File() = Nothing
        Dim oFileVersion As ACW.File
        Dim oLatestFileByLifeCycle As ACW.File = Nothing
        Dim oFileIteration As VDF.Vault.Currency.Entities.FileIteration
        Dim strFullFileName As String
        Dim oServices As VDF.Vault.Services.Connection.IWorkingFoldersManager
        Dim oDocumentService As ACW.DocumentService = oConnection.WebServiceManager.DocumentService

        '*******************************************************************************************
        ' ESTABLISH VAULT CONNECTION USING VAULT ADD-IN WITHIN INVENTOR
        '*******************************************************************************************
        If Vault_IsConnected() = False Then Exit Function

        '*******************************************************************************************
        ' FIND FILE 
        '*******************************************************************************************
        Try
            oFile = File_FindByFilename(strFileName)
            If strFileName = "" Then strFileName = oFile.Name
        Catch ex As Exception
            MsgBox("File was not able to be found, make sure strFileName is correct and Vault Add-in is logged in.", MsgBoxStyle.OkOnly, "Error")
            Exit Function
        End Try

        '*******************************************************************************************
        ' CREATE AN ARRAY OF FILE VERSIONS TO ITERATE THROUGH
        '*******************************************************************************************

        Try
            If oFile IsNot Nothing Then oFileVersions = oDocumentService.GetFilesByMasterId(oFile.MasterId)
        Catch ex As Exception
            MsgBox("Master ID was not available.", MsgBoxStyle.OkOnly, "Error")
            Exit Function
        End Try

        '*******************************************************************************************
        ' REVERSE THE ARRAY TO GO FROM EARLIEST VERSION TO LATEST
        '*******************************************************************************************
        If oFileVersions.Count > 0 Then Array.Reverse(oFileVersions)

        '*******************************************************************************************
        ' FINDS THE LATEST VERSION OF A FILE AT THE LIFECYCLE STATE NEEDED
        '*******************************************************************************************
        'INCREASES BY THE LATEST VERSION FORWARD
        For Each oFileVersion In oFileVersions
            If oFileVersion.FileLfCyc.LfCycStateName.ToUpper = strLifeCycleState Then
                oLatestFileByLifeCycle = oFileVersion
                Exit For
            End If
        Next

        If oLatestFileByLifeCycle Is Nothing Then
            MsgBox("A file at that lifecycle state was not found.", MsgBoxStyle.OkOnly, "Error")
            Exit Function
        End If

        '*******************************************************************************************
        ' TRY TO GET THE FILE FROM VAULT IF NOT CHECKED OUT
        '*******************************************************************************************

        If oLatestFileByLifeCycle IsNot Nothing Then
            oFileIteration = New VDF.Vault.Currency.Entities.FileIteration(oConnection, oFile)
            oServices = oConnection.WorkingFoldersManager
            strFullFileName = oServices.GetPathOfFileInWorkingFolder(oFileIteration).FullPath.ToString

            Try
                If File_IsCheckedOut(strFileName) = False Then
                    File_Acquire(oLatestFileByLifeCycle, False)
                    GetFile_LatestVersionByLifecycleState = strFullFileName
                Else
                    MsgBox("Get was not possible file is checked out.", MsgBoxStyle.OkOnly, "Error")
                    Exit Function
                End If

            Catch ex As Exception
                MsgBox("Get file was not successful.", MsgBoxStyle.OkOnly, "Error")
                Exit Function
            End Try
        End If

    End Function

#End Region

#Region "_09 GetFromFile_LatestVersionLifecycleState"
    Public Function GetFromFile_LatestVersionsLifecycleState(strFileName As String) As String
        '*******************************************************************************************
        ' THIS FUNCTION WILL GET THE LIFECYCLE STATE OF THE LATEST VERSION IN VAULT
        '*******************************************************************************************

        '*******************************************************************************************
        ' CHECK MAJOR REQUIREMENTS FOR THIS FUNCTION AND EXITS EARLY IF NEEDED. 
        '*******************************************************************************************
        GetFromFile_LatestVersionsLifecycleState = ""
        If strFileName = "" Then
            MsgBox("strFileName must have a value to continue", MsgBoxStyle.OkOnly, "Error")
            Exit Function
        End If

        Dim oFileVersions As ACW.File() = Nothing
        Dim oFile As ACW.File
        Dim oDocumentService As ACW.DocumentService

        '*******************************************************************************************
        ' ESTABLISH VAULT CONNECTION USING VAULT ADD-IN WITHIN INVENTOR
        '*******************************************************************************************
        If Vault_IsConnected() = False Then Exit Function

        '*******************************************************************************************
        ' SETTING THE DOCUMENT SERVICE OBJECT
        '*******************************************************************************************
        oDocumentService = oConnection.WebServiceManager.DocumentService

        '*******************************************************************************************
        ' FIND FILE 
        '*******************************************************************************************
        Try
            oFile = File_FindByFilename(strFileName)
        Catch ex As Exception
            MsgBox("File was not able to be found, make sure strFileName is correct and Vault Add-in is logged in.", MsgBoxStyle.OkOnly, "Error")
            Exit Function
        End Try

        '*******************************************************************************************
        ' GET ALL VERSIONS FOR FILE
        '*******************************************************************************************
        Try
            If oFile IsNot Nothing Then oFileVersions = oDocumentService.GetFilesByMasterId(oFile.MasterId)
        Catch ex As Exception
            MsgBox("File versions were not available.", MsgBoxStyle.OkOnly, "Error")
            Exit Function
        End Try

        '*******************************************************************************************
        ' REVERSE THE ARRAY TO GO FROM EARLIEST VERSION TO LATEST
        '*******************************************************************************************
        Try
            If oFileVersions IsNot Nothing Then Array.Reverse(oFileVersions)
        Catch ex As Exception
            MsgBox("Reversing file versions was not possible.", MsgBoxStyle.OkOnly, "Error")
            Exit Function
        End Try

        '*******************************************************************************************
        ' SET THE FIRST INSTANCE OF THE ARRAY AND IT'S LIFECYCLE TO THE FUNCTION
        '*******************************************************************************************
        Try
            If oFileVersions IsNot Nothing Then GetFromFile_LatestVersionsLifecycleState = oFileVersions(0).FileLfCyc.LfCycStateName
        Catch ex As Exception
            MsgBox("Latest Lifecycle state name was not available.", MsgBoxStyle.OkOnly, "Error")
            Exit Function
        End Try

    End Function

#End Region

#Region "_10 LocalFolder_ClearContents"
    Public Sub LocalFolder_ClearContents(strFolderPath As String)
        '*******************************************************************************************
        ' THIS FUNCTION WILL CLEAR ALL FOLDERS AND FILES FROM A GIVEN FOLDER
        '*******************************************************************************************
        Dim strFiles As ObjectModel.ReadOnlyCollection(Of String) = FileIO.FileSystem.GetFiles(strFolderPath)
        Dim strDirectores As ObjectModel.ReadOnlyCollection(Of String) = FileIO.FileSystem.GetDirectories(strFolderPath)
        Dim strFileFolder As String

        '*******************************************************************************************
        ' DELETE ALL FILES IN A FOLDER
        '*******************************************************************************************
        For Each strFileFolder In strFiles
            FileIO.FileSystem.DeleteFile(strFileFolder, FileIO.UIOption.OnlyErrorDialogs, FileIO.RecycleOption.DeletePermanently, FileIO.UICancelOption.DoNothing)
        Next

        '*******************************************************************************************
        ' DELETE ALL FODLERS IN A FOLDER
        '*******************************************************************************************
        For Each strFileFolder In strDirectores
            FileIO.FileSystem.DeleteDirectory(strFileFolder, FileIO.UIOption.OnlyErrorDialogs, FileIO.RecycleOption.DeletePermanently, FileIO.UICancelOption.DoNothing)
        Next


    End Sub

#End Region

#Region "_11 Item_GetPropVal"

    Public Function Item_GetPropVal(strItemNumber As String, strPropNameToFind As String, Optional blnIncludeAll As Boolean = True, Optional blnIncludeInternal As Boolean = False,
                                    Optional blnIncludeSystem As Boolean = False, Optional blnIncludeUserDefined As Boolean = False, Optional blnOnlyActive As Boolean = False) As String

        '*******************************************************************************************
        ' CHECK MAJOR REQUIREMENTS FOR THIS FUNCTION AND EXITS EARLY IF NEEDED. 
        '*******************************************************************************************
        Item_GetPropVal = ""
        If strItemNumber = "" Or strPropNameToFind = "" Then
            MsgBox("Both inputs need a value to return a property value.", MsgBoxStyle.OkOnly, "Error")
            Exit Function
        End If


        Dim oItemService As ACW.ItemService = oConnection.WebServiceManager.ItemService
        Dim oItem = oItemService.GetLatestItemByItemNumber(strItemNumber)
        Dim oProps As PropertyDefinitionDictionary
        Dim oItemRev As New ItemRevision(oConnection, oItem)
        Dim oFilter As PropertyDefinitionFilter
        Dim oPropVal As Object

        '*******************************************************************************************
        ' DEFINE FILTER
        '*******************************************************************************************
        If blnIncludeInternal Then
            oFilter = PropertyDefinitionFilter.IncludeInternal
        ElseIf blnIncludeSystem Then
            oFilter = PropertyDefinitionFilter.IncludeSystem
        ElseIf blnIncludeUserDefined Then
            oFilter = PropertyDefinitionFilter.IncludeUserDefined
        ElseIf blnOnlyActive Then
            oFilter = PropertyDefinitionFilter.OnlyActive
        Else
            oFilter = PropertyDefinitionFilter.IncludeAll
        End If

        '*******************************************************************************************
        ' CREATE PROPERTY DEFINITIONS
        '*******************************************************************************************
        oProps = oConnection.PropertyManager.GetPropertyDefinitions("ITEM", Nothing, oFilter)

        '*******************************************************************************************
        ' LOOP THROUGH PROPERTY DEFINITIONS TO FIND KEY THAT MATCHES INPUT strPropNameToFind
        '*******************************************************************************************
        For Each key In oProps.Keys
            Try
                If oProps(key).DisplayName.ToUpper = strPropNameToFind.ToUpper Then
                    oPropVal = oConnection.PropertyManager.GetPropertyValue(oItemRev, oProps(key), Nothing)
                    If oPropVal IsNot Nothing Then
                        Item_GetPropVal = oPropVal
                        Exit For
                    End If
                End If
            Catch ex As Exception

            End Try
        Next


    End Function

#End Region

#Region "_12 Item_GetAllPropsAndVals"

    Public Function Item_GetAllPropsAndVals(strItemNumber As String) As Dictionary(Of String, String)

        '*******************************************************************************************
        ' CHECK MAJOR REQUIREMENTS FOR THIS FUNCTION AND EXITS EARLY IF NEEDED. 
        '*******************************************************************************************
        Item_GetAllPropsAndVals = Nothing
        If strItemNumber = "" Then
            MsgBox("strItemNumber need a value..", MsgBoxStyle.OkOnly, "Error")
            Exit Function
        End If


        Dim oItemService As ACW.ItemService = oConnection.WebServiceManager.ItemService
        Dim oItem = oItemService.GetLatestItemByItemNumber(strItemNumber)
        Dim oProps As PropertyDefinitionDictionary
        Dim oItemRev As New ItemRevision(oConnection, oItem)
        Dim oFilter As PropertyDefinitionFilter
        Dim oPropVal As Object
        Dim oPropDict As New Dictionary(Of String, String)

        '*******************************************************************************************
        ' DEFINE FILTER
        '*******************************************************************************************

        oFilter = PropertyDefinitionFilter.IncludeAll


        '*******************************************************************************************
        ' CREATE PROPERTY DEFINITIONS
        '*******************************************************************************************
        oProps = oConnection.PropertyManager.GetPropertyDefinitions("ITEM", Nothing, oFilter)

        '*******************************************************************************************
        ' LOOP THROUGH PROPERTY DEFINITIONS TO FIND KEY THAT MATCHES INPUT strPropNameToFind
        '*******************************************************************************************
        For Each key In oProps.Keys
            Try
                oPropVal = oConnection.PropertyManager.GetPropertyValue(oItemRev, oProps(key), Nothing)
                If oPropVal IsNot Nothing Then
                    oPropDict.Add(oProps(key).DisplayName, oPropVal.ToString)
                End If
            Catch ex As Exception

            End Try
        Next

        Item_GetAllPropsAndVals = oPropDict

    End Function

#End Region

#Region "_13 Item_FindAllAssociatedFiles"
    Public Function Item_FindAllAssociatedFiles(strItemNumber As String) As Dictionary(Of String, String)
        Dim oItemService As ACW.ItemService = oConnection.WebServiceManager.ItemService
        Dim oItem = oItemService.GetLatestItemByItemNumber(strItemNumber)
        Dim oFileAssocs As ACW.ItemFileAssoc()
        Dim oFileAssoc As ACW.ItemFileAssoc
        Dim lstFileAssoc As New List(Of String)
        Dim oFileDict As New Dictionary(Of String, String)

        '*******************************************************************************************
        ' FIND PRIMARY ASSOCIATED FILES
        '*******************************************************************************************
        oFileAssocs = oItemService.GetItemFileAssociationsByItemIds({oItem.Id}, ACW.ItemFileLnkTypOpt.Primary)

        If oFileAssocs.Count > 0 Then
            For Each oFileAssoc In oFileAssocs
                Try
                    oFileDict.Add("Primary", oFileAssoc.FileName)
                Catch ex As Exception

                End Try
            Next
        Else
            oFileDict.Add("Primary", "")
        End If

        '*******************************************************************************************
        ' FIND PRIMARY SUB ASSOCIATED FILES
        '*******************************************************************************************
        oFileAssocs = oItemService.GetItemFileAssociationsByItemIds({oItem.Id}, ACW.ItemFileLnkTypOpt.PrimarySub)

        If oFileAssocs.Count > 0 Then
            For Each oFileAssoc In oFileAssocs
                Try
                    oFileDict.Add("PrimarySub", oFileAssoc.FileName)
                Catch ex As Exception

                End Try
            Next
        Else
            oFileDict.Add("PrimarySub", "")
        End If

        '*******************************************************************************************
        ' FIND SECONDARY ASSOCIATED FILES
        '*******************************************************************************************
        oFileAssocs = oItemService.GetItemFileAssociationsByItemIds({oItem.Id}, ACW.ItemFileLnkTypOpt.Secondary)

        If oFileAssocs.Count > 0 Then
            For Each oFileAssoc In oFileAssocs
                Try
                    oFileDict.Add("Secondary", oFileAssoc.FileName)
                Catch ex As Exception

                End Try
            Next
        Else
            oFileDict.Add("Secondary", "")
        End If

        '*******************************************************************************************
        ' FIND SECONDARY SUB ASSOCIATED FILES
        '*******************************************************************************************
        oFileAssocs = oItemService.GetItemFileAssociationsByItemIds({oItem.Id}, ACW.ItemFileLnkTypOpt.SecondarySub)

        If oFileAssocs.Count > 0 Then
            For Each oFileAssoc In oFileAssocs
                Try
                    oFileDict.Add("SecondarySub", oFileAssoc.FileName)
                Catch ex As Exception

                End Try
            Next
        Else
            oFileDict.Add("SecondarySub", "")
        End If

        '*******************************************************************************************
        ' FIND STANDARD COMPONENT ASSOCIATED FILES
        '*******************************************************************************************
        oFileAssocs = oItemService.GetItemFileAssociationsByItemIds({oItem.Id}, ACW.ItemFileLnkTypOpt.StandardComponent)

        If oFileAssocs.Count > 0 Then
            For Each oFileAssoc In oFileAssocs
                Try
                    oFileDict.Add("StandardComponent", oFileAssoc.FileName)
                Catch ex As Exception

                End Try
            Next
        Else
            oFileDict.Add("StandardComponent", "")
        End If

        '*******************************************************************************************
        ' FIND TERTIARY ASSOCIATED FILES
        '*******************************************************************************************
        oFileAssocs = oItemService.GetItemFileAssociationsByItemIds({oItem.Id}, ACW.ItemFileLnkTypOpt.Tertiary)

        If oFileAssocs.Count > 0 Then
            For Each oFileAssoc In oFileAssocs
                Try
                    oFileDict.Add("Tertiary", oFileAssoc.FileName)
                Catch ex As Exception

                End Try
            Next
        Else
            oFileDict.Add("Tertiary", "")
        End If



        Item_FindAllAssociatedFiles = oFileDict


    End Function

#End Region


End Class







