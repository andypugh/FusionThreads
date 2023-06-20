Attribute VB_Name = "Module1"
Dim fso As Object
Public Sub export_xml()
    Name = ActiveSheet.Range("B1").Value
    f = FreeFile
    Open ActiveWorkbook.Path & "/" & Name & ".xml" For Output As #f
    With ActiveSheet
    Set R = .Range("B8:N8")
    Print #f, "<?xml version=""1.0"" encoding = ""UTF-8""?>"
    Print #f, "<ThreadType>"
    Print #f, "  <Name>" & .Range("B1") & "</Name>"
    Print #f, "  <CustomName>" & .Range("B1") & "</CustomName>"
    Print #f, "  <Unit>" & .Range("B2") & "</Unit>"
    Print #f, "  <Angle>" & .Range("B3") & "</Angle>"
    Print #f, "  <SortOrder>" & .Range("B4") & "</SortOrder>"
    If .Range("B5") <> "" Then
        'If not defined, default shape is trapezoid = 0. Others: 1=sharp, 5=square, 7=whitworth
        Print #f, "  <ThreadForm>" & .Range("B5") & "</ThreadForm>"
    End If
    While R(1) <> ""
        DoEvents
        Print #f, "  <ThreadSize>"
        Print #f, "    <Size>" & R(1) & "</Size>"
        Print #f, "    <Designation>"
        Print #f, "      <ThreadDesignation>" & R(2) & "</ThreadDesignation>"
        Print #f, "      <CTD>" & R(3) & "</CTD>"
        'TPI or Pitch are valid tags depending on thread
        If StrComp("TPI", .Range("E7"), vbTextCompare) Then
            Print #f, "      <TPI>" & R(4) & "</TPI>"
        ElseIf StrComp("Pitch", .Range("E7"), vbTextCompare) Then
            Print #f, "      <Pitch>" & R(4) & "</Pitch>"
        End If
        Print #f, "      <Thread>"
        Print #f, "        <Gender>external</Gender>"
        Print #f, "        <Class>" & R(5) & "</Class>"
        Print #f, "        <MajorDia>" & R(6) & "</MajorDia>"
        Print #f, "        <PitchDia>" & R(7) & "</PitchDia>"
        Print #f, "        <MinorDia>" & R(8) & "</MinorDia>"
        Print #f, "      </Thread>"
        Print #f, "      <Thread>"
        Print #f, "        <Gender>internal</Gender>"
        Print #f, "        <Class>" & R(9) & "</Class>"
        Print #f, "        <MajorDia>" & R(10) & "</MajorDia>"
        Print #f, "        <PitchDia>" & R(11) & "</PitchDia>"
        Print #f, "        <MinorDia>" & R(12) & "</MinorDia>"
        If R(13) <> "" Then
            Print #f, "        <TapDrill>" & R(13) & "</TapDrill>"
        End If
        Print #f, "      </Thread>"
        Print #f, "    </Designation>"
        Print #f, "  </ThreadSize>"
        Set R = R.Offset(1, 0)
    Wend
    Print #f, "  </ThreadType>"
    End With
    Close #f
End Sub

