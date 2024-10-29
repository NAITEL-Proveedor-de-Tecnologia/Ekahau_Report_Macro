' Copyright © "NAITEL Proveedor de Tecnología y Redes S DE RL DE CV" 2024
' Esta macro está diseñada para cambiar los estilos de encabezado en informes de Ekahau.

Sub ConvertToHeadings()
    Dim para As Paragraph
    Dim headerText As String
    Dim countH2 As Integer
    Dim countH3 As Integer

    countH2 = 0
    countH3 = 0

    ' Heading 2 prefixes
    Dim h2Patterns As Variant
    h2Patterns = Array("associated access point for", "access points on", _
                       "other access points on")

    ' Heading 3 prefixes
    Dim h3Patterns As Variant
    h3Patterns = Array("signal strength for", "channel interference for", _
                       "data rate for", "round trip time for", _
                       "throughput for", "network issues for", _
                       "network health for", "number of aps for", _
                       "channel utilization for", "packet loss for")

    For Each para In ActiveDocument.Paragraphs
        headerText = LCase(Trim(para.Range.Text))
        
        ' Check for Heading 2 patterns
        For Each Prefix In h2Patterns
            If InStr(headerText, Prefix) = 1 Then
                On Error Resume Next
                para.Style = ActiveDocument.Styles("Título 2")
                If Err.Number <> 0 Then
                    MsgBox "The style 'Título 2' does not exist."
                    Err.Clear
                End If
                On Error GoTo 0
                countH2 = countH2 + 1
                Exit For
            End If
        Next Prefix
        
        ' Check for Heading 3 patterns
        For Each Prefix In h3Patterns
            If InStr(headerText, Prefix) = 1 Then
                On Error Resume Next
                para.Style = ActiveDocument.Styles("Título 3")
                If Err.Number <> 0 Then
                    MsgBox "The style 'Título 3' does not exist."
                    Err.Clear
                End If
                On Error GoTo 0
                countH3 = countH3 + 1
                Exit For
            End If
        Next Prefix
    Next para
    
    MsgBox countH2 & " headings changed to Título 2, " & countH3 & " headings changed to Título 3."
End Sub


