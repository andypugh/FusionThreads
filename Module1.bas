Attribute VB_Name = "Module1"
Dim fso As Object
Public Sub export_xml()
    Name = ActiveSheet.Range("B1").Value
    f = FreeFile
    Open ActiveWorkbook.Path & "/" & Name & ".xml" For Output As #f
    With ActiveSheet
    Set R = .Range("A8:N8")
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
        Print #f, "    <Size>" & R(2) & "</Size>"
        Print #f, "    <Designation>"
        Print #f, "      <ThreadDesignation>" & R(4) & "</ThreadDesignation>"
        Print #f, "      <CTD>" & R(5) & "</CTD>"
        'TPI or Pitch are valid tags depending on thread
        If StrComp("TPI", .Range("C7"), vbTextCompare) Then
            Print #f, "      <TPI>" & R(3) & "</TPI>"
        ElseIf StrComp("Pitch", .Range("C7"), vbTextCompare) Then
            Print #f, "      <Pitch>" & R(3) & "</Pitch>"
        End If
        Print #f, "      <Thread>"
        Print #f, "        <Gender>external</Gender>"
        Print #f, "        <Class>" & R(6) & "</Class>"
        Print #f, "        <MajorDia>" & R(8) & "</MajorDia>"
        Print #f, "        <PitchDia>" & R(9) & "</PitchDia>"
        Print #f, "        <MinorDia>" & R(10) & "</MinorDia>"
        Print #f, "      </Thread>"
        Print #f, "      <Thread>"
        Print #f, "        <Gender>internal</Gender>"
        Print #f, "        <Class>" & R(11) & "</Class>"
        Print #f, "        <MajorDia>" & R(13) & "</MajorDia>"
        Print #f, "        <PitchDia>" & R(14) & "</PitchDia>"
        Print #f, "        <MinorDia>" & R(15) & "</MinorDia>"
        If R(16) <> "" Then
            Print #f, "        <TapDrill>" & R(16) & "</TapDrill>"
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

