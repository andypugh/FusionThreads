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
    Print #f, "  <ThreadForm>" & .Range("B5") & "</ThreadForm>"
    While R(1) <> ""
        DoEvents
        Print #f, "  <ThreadSize>"
        Print #f, "    <Size>" & R(1) & "</Size>"
        Print #f, "    <Designation>"
        Print #f, "      <ThreadDesignation>" & R(2) & "</ThreadDesignation>"
        Print #f, "      <CTD>" & R(3) & "</CTD>"
        Print #f, "      <TPI>" & R(4) & "</TPI>"
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
        Print #f, "      </Thread>"
        Print #f, "    </Designation>"
        Print #f, "  </ThreadSize>"
        Set R = R.Offset(1, 0)
    Wend
    Print #f, "  </ThreadType>"
    End With
    Close #f
End Sub

