Attribute VB_Name = "writeXml"
Sub Click_XmlOutputButton()
    Application.ScreenUpdating = False
    Dim tmppath As String
    Dim outpath As String
    Dim filepath As String
    Dim buf As String, cnt As Long
    
    tmppath = ThisWorkbook.path & "\config_tmp"
    outpath = ThisWorkbook.path & "\config"
    
    Debug.Print tmppath
    Debug.Print outpath
    
    If Dir(tmppath, vbDirectory) = "" Then
        MkDir tmppath
    Else
        buf = Dir(tmppath & "\*.xml?")
        Do While buf <> ""
            Debug.Print buf
            Kill tmppath & "\" & buf
            buf = Dir()
        Loop
        RmDir tmppath
        MkDir tmppath
    End If
    If Dir(outpath, vbDirectory) = "" Then
        MkDir outpath
    Else
        buf = Dir(outpath & "\*.xml?")
        Do While buf <> ""
            Debug.Print buf
            Kill outpath & "\" & buf
            buf = Dir()
        Loop
        RmDir outpath
        MkDir outpath
    End If
    Dim path As String
    For Each w In ThisWorkbook.Worksheets
        w.Select
        filepath = writeXml(tmppath)
        filepath = loadXml(outpath, filepath)
        Debug.Print filepath
    Next
    Application.ScreenUpdating = True
End Sub

Function writeXml(ByVal tmppath As String) As String

    ' Microsoft XML v6.0を使用
    ' 参照設定で「Microsoft XML, v6.0」にチェックを入れて下さい
    Dim xD As New MSXML2.DOMDocument60
    Dim nd(2) As MSXML2.IXMLDOMNode
    Dim rowIndex As Long
    Dim fileName As String
    Dim xmlpath As String
    
    
    ' 親ノード作成
    Set nd(0) = xD.createNode(NODE_ELEMENT, "xmltest", "")
    
    ' 1個目のノード作成
    Set nd(1) = xD.createNode(NODE_ELEMENT, ActiveSheet.Name, "")
    nd(0).appendChild nd(1)     ' 親ノードに1個目のノードを追加
    
    rowIndex = 2
    Do While Cells(rowIndex, 2) <> ""
        Set nd(2) = xD.createNode(NODE_ELEMENT, Cells(rowIndex, 2).Value, "")
        nd(2).Text = Cells(rowIndex, 4).Value
        nd(1).appendChild nd(2)
        rowIndex = rowIndex + 1
    Loop
       
    ' ルート作成
    xD.appendChild xD.createProcessingInstruction("xml", "version='1.0' encoding='UTF-8'")
    xD.appendChild nd(0)      ' ルートに親ノードを追加

    fileName = ActiveSheet.Name & ".xml"
    xmlpath = tmppath & "\" & fileName
    Debug.Print path
    
    ' ファイルに保存
    xD.Save xmlpath
    
    Debug.Print xD.xml
    
    Set xD = Nothing
    
    writeXml = xmlpath
    
End Function

Function loadXml(ByVal outpath As String, ByVal filepath As String) As String

    Dim xD As New MSXML2.DOMDocument60

    outpath = outpath & ActiveSheet.Name & ".xml"
        
    xD.Load indent(filepath)
    
    Debug.Print xD.xml
        
    xD.Save outpath
    
    Set xD = Nothing
End Function

Function indent(ByVal xml As String) As String
    Dim writer As MSXML2.MXXMLWriter60
    Dim reader As MSXML2.SAXXMLReader60
    Dim dom As MSXML2.DOMDocument60
    Dim n As MSXML2.IXMLDOMNode

    Set writer = New MSXML2.MXXMLWriter60
    ' xml宣言を書き込まない
    writer.omitXMLDeclaration = True
    ' インデントする
    writer.indent = True
    
    Set reader = New MSXML2.SAXXMLReader60
    Set reader.contentHandler = writer
    reader.Parse (xml)

    ' 元のxmlから、xml宣言候補を退避
    Set dom = New MSXML2.DOMDocument60
    dom.loadXml xml
    Set n = dom.ChildNodes(0)

    ' インデントされたxmlを読み込む
    ' 元のxmlにxml宣言があったとしても、除外されている
    dom.loadXml writer.output

    ' 元のxmlにxml宣言があれば、インデントされたxmlに追加
    If n.nodeName = "xml" And n.NodeType = NODE_PROCESSING_INSTRUCTION Then
        dom.InsertBefore n, dom.ChildNodes(0)
    End If
    
    Debug.Print dom.xml

    ' インデントされたxmlを返す
    indent = dom.xml
End Function
