Attribute VB_Name = "modMain"
Option Explicit

Global Const LISTVIEW_MODE0 = "View Large Icons"
Global Const LISTVIEW_MODE1 = "View Small Icons"
Global Const LISTVIEW_MODE2 = "View List"
Global Const LISTVIEW_MODE3 = "View Details"
Public fMainForm As frmMain
Public cn As New ADODB.Connection
Const BLOCKSIZE As Long = 4096

Public Sub ColumnToFile(Col As ADODB.Field, _
                        RtrnString As String)
    'Retrieves data from the database and puts it into a temp file on
    'the hard drive.
    'The size of the chunk is in the variable BLOCKSIZE (4096).

    Dim NumBlocks As Long  'Holds the number of chunks.
    Dim LeftOver As Long   '# of chars left over after last whole chunk.
    Dim strData As String
    Dim DestFileNum As Long
    Dim I As Long
    Dim ColSize As Long


        ColSize = Col.ActualSize

        NumBlocks = ColSize \ BLOCKSIZE
        LeftOver = ColSize Mod BLOCKSIZE

        'Now Write data to the file in chunks.
        For I = 1 To NumBlocks
            strData = String(BLOCKSIZE, 0)
            strData = Col.GetChunk(BLOCKSIZE)
            RtrnString = RtrnString & strData
        Next I

        strData = String(LeftOver, 0)
        strData = Col.GetChunk(LeftOver)
        RtrnString = RtrnString & strData

End Sub

Sub FileToColumn(Col As ADODB.Field, _
                 DiskFile As String)
    'Takes data from the temp file and saves it to the database.

    Dim strData As String
    Dim NumBlocks As Long
    Dim FileLength As Long
    Dim LeftOver As Long
    Dim SourceFile As Long
    Dim I As Long
    Const BLOCKSIZE As Long = 4096

    SourceFile = FreeFile
    Open DiskFile For Binary Access Read As SourceFile
    FileLength = LOF(SourceFile)

    If FileLength = 0 Then
        Close SourceFile
        MsgBox DiskFile & " Empty or Not Found."
    Else
        NumBlocks = FileLength \ BLOCKSIZE
        LeftOver = FileLength Mod BLOCKSIZE
        Col.AppendChunk Null
        strData = String(BLOCKSIZE, 0)

        For I = 1 To NumBlocks
            Get SourceFile, , strData
            Col.AppendChunk strData
        Next I

        strData = String(LeftOver, 0)
        Get SourceFile, , strData
        Col.AppendChunk strData
        ' rsset.Update
        Close SourceFile
    End If

End Sub

Sub Main()
    frmSplash.Show
    frmSplash.Refresh
    Set fMainForm = New frmMain
    Load fMainForm


    fMainForm.Show


End Sub

Public Function LoadFeed(strURL As String) As FreeThreadedDOMDocument30

    Dim xmlhttp As New MSXML2.xmlhttp
    Dim Source As New DOMDocument30
    Dim errorXML As String
    
    xmlhttp.Open "GET", strURL, False
    xmlhttp.send

    Set Source = New MSXML2.FreeThreadedDOMDocument40

    Source.async = False
    Source.loadXML (xmlhttp.responseXML.xml)
    errorXML = Source.parseError.errorCode


    
    Set LoadFeed = Source
    Set Source = Nothing

End Function

Public Function WriteOption(OptionName As String, Value As String)
    
    Dim oDom As New FreeThreadedDOMDocument30
    Dim oNode As IXMLDOMNode
    
    oDom.Load App.path & "\settings.xml"
    
    Set oNode = oDom.selectSingleNode("//options/" & OptionName)
    
    If oNode Is Nothing Then
        Set oNode = oDom.createElement("optionName")
        oDom.selectSingleNode("//option").appendChild oNode
    End If
    
    oNode.Text = Value
    
    oDom.Save App.path & "\settings.xml"
        
End Function

Public Function GetOption(OptionName As String, Optional DefaultValue As String = "") As String
    
    Dim oDom As New FreeThreadedDOMDocument30
    Dim oNode As IXMLDOMNode
    
    oDom.Load App.path & "\settings.xml"

    Set oNode = oDom.selectSingleNode("//options/" & OptionName)
    
    If oNode Is Nothing Then
        GetOption = DefaultValue
    Else
        GetOption = oNode.Text
    End If

End Function
