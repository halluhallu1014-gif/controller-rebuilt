Attribute VB_Name = "BuildOfficeController"
Option Explicit

' GitHub公開用。添付の画面からレイアウトのみを再構築し、実データ・実パスは含めない。
' 新規空ブック(.xlsm)に追加してBuildOfficeControllerを実行する。
Public Sub BuildOfficeController()
    Dim book As Workbook
    Dim searchSheet As Worksheet, detailSheet As Worksheet
    Dim addressSheet As Worksheet, fieldsSheet As Worksheet
    Set book = ThisWorkbook
    If book.Worksheets.Count <> 1 Then
        MsgBox "新しく作った空のxlsmブックで実行してください。", vbExclamation: Exit Sub
    End If
    If Application.WorksheetFunction.CountA(book.Worksheets(1).Cells) > 0 Then
        MsgBox "最初のシートにデータがあるため中止します。", vbExclamation: Exit Sub
    End If
    Application.ScreenUpdating = False
    On Error GoTo Failed
    Set searchSheet = book.Worksheets(1): searchSheet.Name = "検索"
    Set detailSheet = book.Worksheets.Add(After:=searchSheet): detailSheet.Name = "詳細"
    Set addressSheet = book.Worksheets.Add(After:=detailSheet): addressSheet.Name = "Address"
    Set fieldsSheet = book.Worksheets.Add(After:=addressSheet): fieldsSheet.Name = "Sheet2"
    BuildSearchLayout searchSheet
    BuildDetailLayout detailSheet
    BuildAddressSettings addressSheet
    BuildFieldSettings fieldsSheet
    searchSheet.Activate
    Application.ScreenUpdating = True
    Debug.Print "CONTROLLER_BUILD=PASS"
    MsgBox "4シートの作成が完了しました。Addressのパスを設定し、業務VBAを追加してください。", vbInformation
    Exit Sub
Failed:
    Application.ScreenUpdating = True
    MsgBox "作成に失敗しました (" & Err.Number & "): " & Err.Description, vbExclamation
End Sub

Private Sub BuildSearchLayout(ByVal sh As Worksheet)
    sh.Cells.Font.Name = "Meiryo UI"
    sh.Cells.Font.Size = 10
    sh.Cells.RowHeight = 18
    sh.Cells.ColumnWidth = 4
    sh.Activate
    sh.Range("A:B").ColumnWidth = 3.125
    sh.Range("C:V").ColumnWidth = 6.000
    sh.Range("W:W").ColumnWidth = 3.125
    sh.Range("X:X").ColumnWidth = 2.000
    sh.Range("Y:XFD").ColumnWidth = 9.000
    sh.Rows(1).RowHeight = 9.00
    sh.Rows(2).RowHeight = 14.10
    sh.Rows(3).RowHeight = 14.10
    sh.Rows(4).RowHeight = 14.10
    sh.Rows(5).RowHeight = 8.10
    sh.Rows(6).RowHeight = 3.95
    sh.Rows(7).RowHeight = 32.10
    sh.Rows(8).RowHeight = 6.00
    sh.Rows(9).RowHeight = 15.00
    sh.Rows(10).RowHeight = 15.00
    sh.Rows(11).RowHeight = 6.00
    sh.Rows(12).RowHeight = 20.10
    sh.Rows(13).RowHeight = 20.10
    sh.Rows(14).RowHeight = 21.95
    sh.Rows(15).RowHeight = 3.95
    sh.Rows(16).RowHeight = 3.95
    sh.Rows(17).RowHeight = 27.95
    sh.Rows(18).RowHeight = 2.10
    sh.Rows(19).RowHeight = 24.00
    sh.Rows(20).RowHeight = 21.00
    sh.Rows(21).RowHeight = 21.00
    sh.Rows(22).RowHeight = 21.00
    sh.Rows(23).RowHeight = 21.00
    sh.Rows(24).RowHeight = 21.00
    sh.Rows(25).RowHeight = 21.00
    sh.Rows(26).RowHeight = 21.00
    sh.Rows(27).RowHeight = 21.00
    sh.Rows(28).RowHeight = 21.00
    sh.Rows(29).RowHeight = 21.00
    sh.Rows(30).RowHeight = 21.00
    sh.Rows(31).RowHeight = 21.00
    sh.Rows(32).RowHeight = 21.00
    sh.Rows(33).RowHeight = 21.00
    sh.Rows(34).RowHeight = 6.00
    sh.Rows(35).RowHeight = 9.95
    sh.Rows(36).RowHeight = 9.95
    sh.Rows(37).RowHeight = 9.95
    sh.Rows(38).RowHeight = 9.95
    sh.Rows(39).RowHeight = 17.10
    sh.Rows(40).RowHeight = 30.00
    sh.Rows(41).RowHeight = 5.10
    sh.Rows(42).RowHeight = 5.10
    sh.Rows(43).RowHeight = 18.75
    sh.Rows(44).RowHeight = 18.75
    sh.Rows(45).RowHeight = 18.75
    sh.Rows(46).RowHeight = 18.75
    sh.Rows(47).RowHeight = 18.75
    sh.Rows(48).RowHeight = 18.75
    sh.Rows(49).RowHeight = 18.75
    sh.Rows(50).RowHeight = 18.75
    sh.Rows(51).RowHeight = 18.75
    sh.Rows(52).RowHeight = 18.75
    sh.Rows(53).RowHeight = 18.75
    sh.Rows(54).RowHeight = 18.75
    sh.Range("T33:U33").Merge
    sh.Range("C35:P37").Merge
    sh.Range("C31:E31").Merge
    sh.Range("F31:J31").Merge
    sh.Range("K31:O31").Merge
    sh.Range("P31:S31").Merge
    sh.Range("C33:E33").Merge
    sh.Range("F33:J33").Merge
    sh.Range("K33:O33").Merge
    sh.Range("P33:S33").Merge
    sh.Range("T31:U31").Merge
    sh.Range("C32:E32").Merge
    sh.Range("F32:J32").Merge
    sh.Range("K32:O32").Merge
    sh.Range("P32:S32").Merge
    sh.Range("T32:U32").Merge
    sh.Range("C29:E29").Merge
    sh.Range("F29:J29").Merge
    sh.Range("K29:O29").Merge
    sh.Range("P29:S29").Merge
    sh.Range("T29:U29").Merge
    sh.Range("C30:E30").Merge
    sh.Range("F30:J30").Merge
    sh.Range("K30:O30").Merge
    sh.Range("P30:S30").Merge
    sh.Range("T30:U30").Merge
    sh.Range("C27:E27").Merge
    sh.Range("F27:J27").Merge
    sh.Range("K27:O27").Merge
    sh.Range("P27:S27").Merge
    sh.Range("T27:U27").Merge
    sh.Range("C28:E28").Merge
    sh.Range("F28:J28").Merge
    sh.Range("K28:O28").Merge
    sh.Range("P28:S28").Merge
    sh.Range("T28:U28").Merge
    sh.Range("C25:E25").Merge
    sh.Range("F25:J25").Merge
    sh.Range("K25:O25").Merge
    sh.Range("P25:S25").Merge
    sh.Range("T25:U25").Merge
    sh.Range("C26:E26").Merge
    sh.Range("F26:J26").Merge
    sh.Range("K26:O26").Merge
    sh.Range("P26:S26").Merge
    sh.Range("T26:U26").Merge
    sh.Range("C23:E23").Merge
    sh.Range("F23:J23").Merge
    sh.Range("K23:O23").Merge
    sh.Range("P23:S23").Merge
    sh.Range("T23:U23").Merge
    sh.Range("C24:E24").Merge
    sh.Range("F24:J24").Merge
    sh.Range("K24:O24").Merge
    sh.Range("P24:S24").Merge
    sh.Range("T24:U24").Merge
    sh.Range("C21:E21").Merge
    sh.Range("F21:J21").Merge
    sh.Range("K21:O21").Merge
    sh.Range("P21:S21").Merge
    sh.Range("T21:U21").Merge
    sh.Range("C22:E22").Merge
    sh.Range("F22:J22").Merge
    sh.Range("K22:O22").Merge
    sh.Range("P22:S22").Merge
    sh.Range("T22:U22").Merge
    sh.Range("T19:U19").Merge
    sh.Range("C20:E20").Merge
    sh.Range("F20:J20").Merge
    sh.Range("K20:O20").Merge
    sh.Range("P20:S20").Merge
    sh.Range("T20:U20").Merge
    sh.Range("P19:S19").Merge
    sh.Range("C14:L14").Merge
    sh.Range("C17:O18").Merge
    sh.Range("C19:E19").Merge
    sh.Range("F19:J19").Merge
    sh.Range("K19:O19").Merge
    sh.Range("B2:W4").Merge
    sh.Range("B5:W5").Merge
    sh.Range("C7:O7").Merge
    sh.Range("C9:D10").Merge
    sh.Range("E9:P10").Merge
    sh.Range("C12:E12").Merge
    sh.Range("F12:I12").Merge
    sh.Range("J12:M12").Merge
    sh.Range("N12:Q12").Merge
    sh.Range("R12:U12").Merge
    ApplyTemplateStyle sh.Range("P7:U7"), 17
    ApplyTemplateStyle sh.Range("B6:W6,B7,V7:W7,B8:W8,B9:B10,Q9:W10,B11:W11,B12,V12:W12,B13:W13,B14,M14:W14,B15:W16,B17:B33,P17:W18,V19:W33,B34:W34,B35:B37,Q35:W37,B38:W38,B39,D39:W39,B40:W42"), 18
    ApplyTemplateStyle sh.Range("A1:X1,A2:A42,X2:X18,X34:X42,A43:X54"), 55
    ApplyTemplateStyle sh.Range("X19"), 56
    ApplyTemplateStyle sh.Range("X20,X22,X24,X26,X28,X30,X32"), 57
    ApplyTemplateStyle sh.Range("X21,X23,X25,X27,X29,X31,X33"), 58
    ApplyTemplateStyle sh.Range("C39"), 59
    ApplyTemplateStyle sh.Range("C9:D10,C12:E12"), 61
    ApplyTemplateStyle sh.Range("B5:W5,F12:U12,C14:L14"), 62
    ApplyTemplateStyle sh.Range("B2:W4"), 63
    ApplyTemplateStyle sh.Range("C7:O7"), 64
    ApplyTemplateStyle sh.Range("E9:P10"), 65
    ApplyTemplateStyle sh.Range("C17:O18"), 66
    ApplyTemplateStyle sh.Range("C19:U19"), 67
    ApplyTemplateStyle sh.Range("C20:U20,C22:U22,C24:U24,C26:U26,C28:U28,C30:U30,C32:U32"), 68
    ApplyTemplateStyle sh.Range("C21:U21,C23:U23,C25:U25,C27:U27,C29:U29,C31:U31,C33:U33"), 69
    ApplyTemplateStyle sh.Range("C35"), 70
    ApplyTemplateStyle sh.Range("D35:P35,C36:P37"), 71
    sh.Range("B2").Value2 = "CONTROLLER / 検索"
    sh.Range("C7").Value2 = "氏名・ふりがなで検索"
    sh.Range("C9").Value2 = "氏名"
    sh.Range("C12").Value2 = "署"
    sh.Range("C19").Value2 = "署"
    sh.Range("F19").Value2 = "氏名"
    sh.Range("K19").Value2 = "ふりがな"
    sh.Range("P19").Value2 = "文書名"
    sh.Range("T19").Value2 = "状態"
    sh.Range("C39").Value2 = "補助操作"
    sh.Range("E9").MergeArea.ClearContents
    sh.Range("C17").Value2 = "検索結果 0件"
    AddCommandButton sh, "E12", "検索", "SearchCases"
    AddCommandButton sh, "C36", "詳細を開く", "ShowSelectedCaseDetail"
    AddCommandButton sh, "C39", "前へ", "SearchPreviousPage"
    AddCommandButton sh, "J39", "次へ", "SearchNextPage"
    AddCommandButton sh, "C43", "新規取込", "ImportWeeklyLedger"
    AddCommandButton sh, "J43", "署別転記", "TransferToFiveOffices"
End Sub

Private Sub BuildDetailLayout(ByVal sh As Worksheet)
    sh.Cells.Font.Name = "Meiryo UI"
    sh.Cells.Font.Size = 10
    sh.Cells.RowHeight = 18
    sh.Cells.ColumnWidth = 4
    sh.Activate
    sh.Range("A:B").ColumnWidth = 3.125
    sh.Range("C:U").ColumnWidth = 6.000
    sh.Range("V:V").ColumnWidth = 12.000
    sh.Range("W:W").ColumnWidth = 3.125
    sh.Range("X:Y").ColumnWidth = 1.500
    sh.Range("Z:AY").ColumnWidth = 9.000
    sh.Range("AZ:AZ").ColumnWidth = 0.000
    sh.Range("BA:XFD").ColumnWidth = 9.000
    sh.Rows(1).RowHeight = 9.00
    sh.Rows(2).RowHeight = 14.10
    sh.Rows(3).RowHeight = 14.10
    sh.Rows(4).RowHeight = 14.10
    sh.Rows(5).RowHeight = 8.10
    sh.Rows(6).RowHeight = 30.95
    sh.Rows(7).RowHeight = 12.00
    sh.Rows(8).RowHeight = 26.10
    sh.Rows(9).RowHeight = 17.10
    sh.Rows(10).RowHeight = 36.95
    sh.Rows(11).RowHeight = 42.95
    sh.Rows(12).RowHeight = 23.10
    sh.Rows(13).RowHeight = 5.10
    sh.Rows(14).RowHeight = 12.95
    sh.Rows(15).RowHeight = 12.95
    sh.Rows(16).RowHeight = 5.10
    sh.Rows(17).RowHeight = 12.95
    sh.Rows(18).RowHeight = 12.95
    sh.Rows(19).RowHeight = 5.10
    sh.Rows(20).RowHeight = 12.95
    sh.Rows(21).RowHeight = 12.95
    sh.Rows(22).RowHeight = 5.10
    sh.Rows(23).RowHeight = 26.10
    sh.Rows(24).RowHeight = 26.10
    sh.Rows(25).RowHeight = 15.00
    sh.Rows(26).RowHeight = 23.10
    sh.Rows(27).RowHeight = 26.10
    sh.Rows(28).RowHeight = 5.10
    sh.Rows(29).RowHeight = 12.95
    sh.Rows(30).RowHeight = 12.95
    sh.Rows(31).RowHeight = 5.10
    sh.Rows(32).RowHeight = 26.10
    sh.Rows(33).RowHeight = 5.10
    sh.Rows(34).RowHeight = 8.10
    sh.Rows(35).RowHeight = 23.10
    sh.Rows(36).RowHeight = 26.10
    sh.Rows(37).RowHeight = 5.10
    sh.Rows(38).RowHeight = 12.95
    sh.Rows(39).RowHeight = 12.95
    sh.Rows(40).RowHeight = 5.10
    sh.Rows(41).RowHeight = 26.10
    sh.Rows(42).RowHeight = 5.10
    sh.Rows(43).RowHeight = 26.10
    sh.Rows(44).RowHeight = 5.10
    sh.Rows(45).RowHeight = 26.10
    sh.Rows(46).RowHeight = 19.50
    sh.Rows(47).RowHeight = 18.75
    sh.Rows(48).RowHeight = 26.10
    sh.Rows(49).RowHeight = 21.95
    sh.Rows(50).RowHeight = 18.95
    sh.Rows(51).RowHeight = 2.10
    sh.Rows(52).RowHeight = 2.10
    sh.Rows(53).RowHeight = 30.00
    sh.Rows(54).RowHeight = 30.00
    sh.Rows(55).RowHeight = 30.00
    sh.Rows(56).RowHeight = 8.10
    sh.Rows(57).RowHeight = 18.00
    sh.Rows(58).RowHeight = 33.00
    sh.Rows(59).RowHeight = 5.10
    sh.Rows(60).RowHeight = 18.95
    sh.Rows(61).RowHeight = 9.95
    sh.Rows(62).RowHeight = 9.95
    sh.Rows(63).RowHeight = 9.95
    sh.Rows(64).RowHeight = 9.95
    sh.Rows(65).RowHeight = 18.00
    sh.Rows(66).RowHeight = 33.00
    sh.Rows(67).RowHeight = 5.10
    sh.Rows(68).RowHeight = 18.95
    sh.Rows(69).RowHeight = 9.95
    sh.Rows(70).RowHeight = 9.95
    sh.Rows(71).RowHeight = 9.95
    sh.Rows(72).RowHeight = 9.95
    sh.Rows(73).RowHeight = 24.95
    sh.Rows(74).RowHeight = 33.00
    sh.Rows(75).RowHeight = 33.00
    sh.Rows(76).RowHeight = 33.00
    sh.Rows(77).RowHeight = 18.00
    sh.Rows(78).RowHeight = 26.10
    sh.Rows(79).RowHeight = 15.95
    sh.Rows(80).RowHeight = 26.10
    sh.Rows(81).RowHeight = 15.95
    sh.Rows(82).RowHeight = 26.10
    sh.Rows(83).RowHeight = 15.95
    sh.Rows(84).RowHeight = 15.95
    sh.Rows(85).RowHeight = 42.95
    sh.Rows(86).RowHeight = 6.00
    sh.Rows(87).RowHeight = 20.10
    sh.Rows(88).RowHeight = 26.10
    sh.Rows(89).RowHeight = 15.95
    sh.Rows(90).RowHeight = 42.95
    sh.Rows(91).RowHeight = 6.00
    sh.Rows(92).RowHeight = 14.25
    sh.Rows(93).RowHeight = 14.25
    sh.Range("C68:V68").Merge
    sh.Range("C69:V72").Merge
    sh.Range("F74:V74").Merge
    sh.Range("F75:V75").Merge
    sh.Range("F76:V76").Merge
    sh.Range("C61:V64").Merge
    sh.Range("D53:L53").Merge
    sh.Range("O53:P53").Merge
    sh.Range("R53:T53").Merge
    sh.Range("U53:V53").Merge
    sh.Range("D54:L54").Merge
    sh.Range("O54:P54").Merge
    sh.Range("R54:T54").Merge
    sh.Range("U54:V54").Merge
    sh.Range("D55:L55").Merge
    sh.Range("O55:P55").Merge
    sh.Range("R55:T55").Merge
    sh.Range("U55:V55").Merge
    sh.Range("C60:V60").Merge
    sh.Range("C48:E48").Merge
    sh.Range("C49:L49").Merge
    sh.Range("O49:Q49").Merge
    sh.Range("R49:T49").Merge
    sh.Range("U49:V49").Merge
    sh.Range("F48:L48").Merge
    sh.Range("C46:E46").Merge
    sh.Range("F46:L46").Merge
    sh.Range("N46:P46").Merge
    sh.Range("Q46:V46").Merge
    sh.Range("C47:E47").Merge
    sh.Range("F47:L47").Merge
    sh.Range("N47:P47").Merge
    sh.Range("Q47:V47").Merge
    sh.Range("C43:E43").Merge
    sh.Range("C45:E45").Merge
    sh.Range("F45:L45").Merge
    sh.Range("N45:P45").Merge
    sh.Range("Q45:V45").Merge
    sh.Range("F43:V43").Merge
    sh.Range("C38:E39").Merge
    sh.Range("F38:L39").Merge
    sh.Range("N38:P39").Merge
    sh.Range("Q38:V39").Merge
    sh.Range("C41:E41").Merge
    sh.Range("F41:L41").Merge
    sh.Range("N41:P41").Merge
    sh.Range("Q41:V41").Merge
    sh.Range("C35:Q35").Merge
    sh.Range("R35:V35").Merge
    sh.Range("C36:E36").Merge
    sh.Range("F36:L36").Merge
    sh.Range("N36:P36").Merge
    sh.Range("Q36:V36").Merge
    sh.Range("C29:E30").Merge
    sh.Range("F29:L30").Merge
    sh.Range("N29:P30").Merge
    sh.Range("Q29:V30").Merge
    sh.Range("C32:E32").Merge
    sh.Range("F32:L32").Merge
    sh.Range("N32:P32").Merge
    sh.Range("Q32:V32").Merge
    sh.Range("C26:Q26").Merge
    sh.Range("R26:V26").Merge
    sh.Range("C27:E27").Merge
    sh.Range("F27:L27").Merge
    sh.Range("N27:P27").Merge
    sh.Range("Q27:V27").Merge
    sh.Range("C20:E21").Merge
    sh.Range("F20:L21").Merge
    sh.Range("N20:P21").Merge
    sh.Range("Q20:V21").Merge
    sh.Range("C23:E24").Merge
    sh.Range("F23:V24").Merge
    sh.Range("C14:E15").Merge
    sh.Range("F14:L15").Merge
    sh.Range("N14:P15").Merge
    sh.Range("Q14:V15").Merge
    sh.Range("C17:E18").Merge
    sh.Range("F17:L18").Merge
    sh.Range("C12:Q12").Merge
    sh.Range("R12:V12").Merge
    sh.Range("B2:W4").Merge
    sh.Range("C6:I6").Merge
    sh.Range("J6:O6").Merge
    sh.Range("P6:S6").Merge
    sh.Range("T6:V6").Merge
    sh.Range("C8:O8").Merge
    sh.Range("Q8:V8").Merge
    sh.Range("C9:D9").Merge
    sh.Range("E9:F9").Merge
    sh.Range("G9:K9").Merge
    sh.Range("L9:O9").Merge
    sh.Range("Q9:V10").Merge
    ApplyTemplateStyle sh.Range("A1:Y1,A2:A47,X2:Y47,A49:A58,X49:Y58,A65:A66,X65:Y66"), 15
    ApplyTemplateStyle sh.Range("B5:W5,C7:R7,C13:Q13,M15,M17:V18,M27,M32,C33:U34,M36,C37:U37,M41,C42:U42"), 17
    ApplyTemplateStyle sh.Range("B6:B7,W6,V7:W7,B9,W9,B11:B15,W11:W12,R13:W13,M14,W14:W15,B16:W16,B17:B18,W17:W18,B19:W19,B20:B21,W20:W21,B22:W22,B23:B24,W23:W24,B25:W25,B26:B27,W26:W27,B28:W28,B29:B30,W29:W30"), 18
    ApplyTemplateStyle sh.Range("B31:W31,B32:B39,W32,V33:W34,W35:W36,V37:W37,M38:M39,W38:W39,B40:W40,B41:B47,W41,V42:W42,W43:W47,B49:B58,W49:W58,B65:B66,W65:W66"), 18
    ApplyTemplateStyle sh.Range("S7:U7"), 19
    ApplyTemplateStyle sh.Range("B8"), 20
    ApplyTemplateStyle sh.Range("P8,W8"), 21
    ApplyTemplateStyle sh.Range("P9"), 22
    ApplyTemplateStyle sh.Range("B10,P10,W10"), 23
    ApplyTemplateStyle sh.Range("C10:O10"), 24
    ApplyTemplateStyle sh.Range("C11:Q11"), 25
    ApplyTemplateStyle sh.Range("R11:V11"), 26
    ApplyTemplateStyle sh.Range("M20:M21,M29:M30"), 27
    ApplyTemplateStyle sh.Range("C44:V44,M45,M48:Y48"), 28
    ApplyTemplateStyle sh.Range("M49"), 29
    ApplyTemplateStyle sh.Range("N49"), 30
    ApplyTemplateStyle sh.Range("C50:C52"), 31
    ApplyTemplateStyle sh.Range("D50:K52"), 32
    ApplyTemplateStyle sh.Range("L50:L52"), 33
    ApplyTemplateStyle sh.Range("M50:M55,O51:U52,Q53:Q54"), 34
    ApplyTemplateStyle sh.Range("N50"), 35
    ApplyTemplateStyle sh.Range("O50:U50"), 36
    ApplyTemplateStyle sh.Range("V50"), 37
    ApplyTemplateStyle sh.Range("N51:N52"), 38
    ApplyTemplateStyle sh.Range("V51:V52"), 39
    ApplyTemplateStyle sh.Range("C53:C54,N53:N54"), 40
    ApplyTemplateStyle sh.Range("C55,N55"), 41
    ApplyTemplateStyle sh.Range("Q55"), 42
    ApplyTemplateStyle sh.Range("C56:V56"), 43
    ApplyTemplateStyle sh.Range("C57,C65"), 44
    ApplyTemplateStyle sh.Range("D57:U57,D65:U65"), 45
    ApplyTemplateStyle sh.Range("V57,V65"), 46
    ApplyTemplateStyle sh.Range("C58,C66"), 47
    ApplyTemplateStyle sh.Range("D58:U58,D66:U66"), 48
    ApplyTemplateStyle sh.Range("V58,V66"), 49
    ApplyTemplateStyle sh.Range("B59:B64,W59:W64,B67:B72,W67:W72,B73:W73"), 50
    ApplyTemplateStyle sh.Range("C59,C67"), 51
    ApplyTemplateStyle sh.Range("D59:U59,D67:U67"), 52
    ApplyTemplateStyle sh.Range("V59,V67"), 53
    ApplyTemplateStyle sh.Range("AZ108:AZ110"), 54
    ApplyTemplateStyle sh.Range("N21:P21,F24:L24,D43:E43,G43:J43"), 62
    ApplyTemplateStyle sh.Range("B2:W4"), 63
    ApplyTemplateStyle sh.Range("C8:O8,C12:Q12,C26:Q26,C35:Q35"), 72
    ApplyTemplateStyle sh.Range("R12:V12,R26:V26,R35:V35"), 73
    ApplyTemplateStyle sh.Range("C6:J6"), 74
    ApplyTemplateStyle sh.Range("K6:S6"), 75
    ApplyTemplateStyle sh.Range("T6:V6"), 76
    ApplyTemplateStyle sh.Range("Q8:V8"), 77
    ApplyTemplateStyle sh.Range("C9:O9"), 78
    ApplyTemplateStyle sh.Range("Q9"), 79
    ApplyTemplateStyle sh.Range("R9:U9"), 80
    ApplyTemplateStyle sh.Range("V9"), 81
    ApplyTemplateStyle sh.Range("Q10"), 82
    ApplyTemplateStyle sh.Range("R10:U10"), 83
    ApplyTemplateStyle sh.Range("V10"), 84
    ApplyTemplateStyle sh.Range("C14:E15,N14:P15,C17:E18,C20:E21,N20:P20,C23,C27:E27,N27:P27,C29:E30,N29,C32:E32,N32:P32,C36:E36,N36:P36,C38:E39,N38:P39,C41:E41,N41,C43,C45:C48,N45:N46"), 85
    ApplyTemplateStyle sh.Range("F14,Q14,F17,F20,Q20,F27,Q27,F29,Q29,F32,Q32,F36,Q36,F38,Q38,F41,Q41,F45:F47,Q45:Q46"), 86
    ApplyTemplateStyle sh.Range("G14:K14,R14:U14"), 87
    ApplyTemplateStyle sh.Range("L14,V14"), 88
    ApplyTemplateStyle sh.Range("F15,Q15"), 89
    ApplyTemplateStyle sh.Range("G15:K15,R15:U15"), 90
    ApplyTemplateStyle sh.Range("L15,V15"), 91
    ApplyTemplateStyle sh.Range("G17:K17"), 92
    ApplyTemplateStyle sh.Range("L17"), 93
    ApplyTemplateStyle sh.Range("F18"), 94
    ApplyTemplateStyle sh.Range("G18:K18"), 95
    ApplyTemplateStyle sh.Range("L18"), 96
    ApplyTemplateStyle sh.Range("G20:K20"), 97
    ApplyTemplateStyle sh.Range("L20"), 98
    ApplyTemplateStyle sh.Range("F21"), 99
    ApplyTemplateStyle sh.Range("G21:K21"), 100
    ApplyTemplateStyle sh.Range("L21"), 101
    ApplyTemplateStyle sh.Range("R20:U20"), 102
    ApplyTemplateStyle sh.Range("V20"), 103
    ApplyTemplateStyle sh.Range("Q21"), 104
    ApplyTemplateStyle sh.Range("R21:U21"), 105
    ApplyTemplateStyle sh.Range("V21"), 106
    ApplyTemplateStyle sh.Range("D23:E23,C24:E24"), 107
    ApplyTemplateStyle sh.Range("F23"), 108
    ApplyTemplateStyle sh.Range("G23:L23"), 109
    ApplyTemplateStyle sh.Range("M23:V24"), 110
    ApplyTemplateStyle sh.Range("G27:K27,R27:U27,G36:K36,R36:U36"), 111
    ApplyTemplateStyle sh.Range("L27,V27,L36,V36"), 112
    ApplyTemplateStyle sh.Range("G29:K29"), 113
    ApplyTemplateStyle sh.Range("L29"), 114
    ApplyTemplateStyle sh.Range("F30"), 115
    ApplyTemplateStyle sh.Range("G30:K30"), 116
    ApplyTemplateStyle sh.Range("L30"), 117
    ApplyTemplateStyle sh.Range("O29:P29,N30:P30"), 118
    ApplyTemplateStyle sh.Range("R29:V29,Q30:V30"), 119
    ApplyTemplateStyle sh.Range("G32:K32,G41:K41"), 120
    ApplyTemplateStyle sh.Range("L32,V32,L41"), 121
    ApplyTemplateStyle sh.Range("R32:U32"), 122
    ApplyTemplateStyle sh.Range("G38:K38"), 123
    ApplyTemplateStyle sh.Range("L38"), 124
    ApplyTemplateStyle sh.Range("F39"), 125
    ApplyTemplateStyle sh.Range("G39:K39"), 126
    ApplyTemplateStyle sh.Range("L39"), 127
    ApplyTemplateStyle sh.Range("R38:V38,Q39:V39"), 128
    ApplyTemplateStyle sh.Range("O41:P41,R41:V41,O45:P45,R45:V45"), 129
    ApplyTemplateStyle sh.Range("D45:E45,D48:E48"), 130
    ApplyTemplateStyle sh.Range("G45:K45"), 131
    ApplyTemplateStyle sh.Range("L45"), 132
    ApplyTemplateStyle sh.Range("F43"), 133
    ApplyTemplateStyle sh.Range("K43"), 134
    ApplyTemplateStyle sh.Range("L43:V43"), 135
    ApplyTemplateStyle sh.Range("D46:E47,G46:L47,O46:P46,R46:V46,N47:V47"), 136
    ApplyTemplateStyle sh.Range("C49"), 137
    ApplyTemplateStyle sh.Range("D49:K49"), 138
    ApplyTemplateStyle sh.Range("L49"), 139
    ApplyTemplateStyle sh.Range("O49"), 140
    ApplyTemplateStyle sh.Range("P49:Q49"), 141
    ApplyTemplateStyle sh.Range("R49:U49"), 142
    ApplyTemplateStyle sh.Range("V49"), 143
    ApplyTemplateStyle sh.Range("F48"), 144
    ApplyTemplateStyle sh.Range("G48:L48"), 145
    ApplyTemplateStyle sh.Range("C61,C69"), 146
    ApplyTemplateStyle sh.Range("D61:U61,D69:U69"), 147
    ApplyTemplateStyle sh.Range("V61,V69"), 148
    ApplyTemplateStyle sh.Range("C62:C63,C70:C71"), 149
    ApplyTemplateStyle sh.Range("D62:U63,D70:U71"), 150
    ApplyTemplateStyle sh.Range("V62:V63,V70:V71"), 151
    ApplyTemplateStyle sh.Range("C64,C72"), 152
    ApplyTemplateStyle sh.Range("D64:U64,D72:U72"), 153
    ApplyTemplateStyle sh.Range("V64,V72"), 154
    ApplyTemplateStyle sh.Range("D53:D54"), 155
    ApplyTemplateStyle sh.Range("E53:K54"), 156
    ApplyTemplateStyle sh.Range("L53:L54"), 157
    ApplyTemplateStyle sh.Range("O53:O54"), 158
    ApplyTemplateStyle sh.Range("P53:P54,R53:T54"), 159
    ApplyTemplateStyle sh.Range("U53:U54"), 160
    ApplyTemplateStyle sh.Range("V53:V54"), 161
    ApplyTemplateStyle sh.Range("D55"), 162
    ApplyTemplateStyle sh.Range("E55:K55"), 163
    ApplyTemplateStyle sh.Range("L55"), 164
    ApplyTemplateStyle sh.Range("O55"), 165
    ApplyTemplateStyle sh.Range("P55,R55:T55"), 166
    ApplyTemplateStyle sh.Range("U55"), 167
    ApplyTemplateStyle sh.Range("V55"), 168
    ApplyTemplateStyle sh.Range("C60,C68"), 169
    ApplyTemplateStyle sh.Range("D60:V60,D68:V68"), 170
    ApplyTemplateStyle sh.Range("F74:F76"), 171
    ApplyTemplateStyle sh.Range("G74:V76"), 172
    sh.Range("B2").Value2 = "CONTROLLER / 事案詳細"
    sh.Range("C9").Value2 = "年度"
    sh.Range("E9").Value2 = "共通番号"
    sh.Range("G9").Value2 = "文書名"
    sh.Range("L9").Value2 = "処理状態"
    sh.Range("C12").Value2 = "01  基本情報"
    sh.Range("R12").Value2 = "本人・連絡先"
    sh.Range("C14").Value2 = "署名"
    sh.Range("N14").Value2 = "氏名法人名"
    sh.Range("C17").Value2 = "共通番号枝番号"
    sh.Range("C20").Value2 = "郵便番号"
    sh.Range("N20").Value2 = "電話番号"
    sh.Range("C23").Value2 = "納税地"
    sh.Range("C26").Value2 = "02  事案内容"
    sh.Range("R26").Value2 = "対象となる申請・請求等"
    sh.Range("C27").Value2 = "税目"
    sh.Range("N27").Value2 = "年分・事業年度等"
    sh.Range("C29").Value2 = "申請書種別"
    sh.Range("N29").Value2 = "提出形式"
    sh.Range("C32").Value2 = "申請日"
    sh.Range("N32").Value2 = "申請又は請求の理由"
    sh.Range("C35").Value2 = "03  処理"
    sh.Range("R35").Value2 = "処理の状態と経過"
    sh.Range("C36").Value2 = "署引継日"
    sh.Range("N36").Value2 = "署引継理由"
    sh.Range("C38").Value2 = "取下書収受日"
    sh.Range("N38").Value2 = "処理年月日"
    sh.Range("C41").Value2 = "施行年月日"
    sh.Range("N41").Value2 = "電子通知"
    sh.Range("C43").Value2 = "摘要"
    sh.Range("C45").Value2 = "業務センター"
    sh.Range("N45").Value2 = "元ファイル"
    sh.Range("C46").Value2 = "処理期限"
    sh.Range("N46").Value2 = "処理担当者"
    sh.Range("C47").Value2 = "処理区分"
    sh.Range("C48").Value2 = "処理グループ（部門）"
    sh.Range("C49").Value2 = "メモ"
    sh.Range("N49").Value2 = "No."
    sh.Range("O49").Value2 = "発送文書・発送日"
    sh.Range("R49").Value2 = "文書名"
    sh.Range("U49").Value2 = "備考"
    sh.Range("C53").Value2 = "メモ1"
    sh.Range("N53").Value2 = "発送1"
    sh.Range("C54").Value2 = "メモ2"
    sh.Range("N54").Value2 = "発送2"
    sh.Range("C55").Value2 = "メモ3"
    sh.Range("N55").Value2 = "発送3"
    sh.Range("AZ108").Value2 = "架空発送文書1"
    sh.Range("AZ109").Value2 = "架空発送文書2"
    sh.Range("AZ110").Value2 = "架空発送文書3"
    sh.Range("Q9").MergeArea.ClearContents
    sh.Range("D53").MergeArea.ClearContents
    sh.Range("D54").MergeArea.ClearContents
    sh.Range("D55").MergeArea.ClearContents
    sh.Range("F14").MergeArea.ClearContents
    sh.Range("F17").MergeArea.ClearContents
    sh.Range("F20").MergeArea.ClearContents
    sh.Range("F23").MergeArea.ClearContents
    sh.Range("F27").MergeArea.ClearContents
    sh.Range("F29").MergeArea.ClearContents
    sh.Range("F32").MergeArea.ClearContents
    sh.Range("F36").MergeArea.ClearContents
    sh.Range("F38").MergeArea.ClearContents
    sh.Range("F41").MergeArea.ClearContents
    sh.Range("F43").MergeArea.ClearContents
    sh.Range("F45").MergeArea.ClearContents
    sh.Range("F46").MergeArea.ClearContents
    sh.Range("F47").MergeArea.ClearContents
    sh.Range("F48").MergeArea.ClearContents
    sh.Range("Q14").MergeArea.ClearContents
    sh.Range("Q20").MergeArea.ClearContents
    sh.Range("Q27").MergeArea.ClearContents
    sh.Range("Q29").MergeArea.ClearContents
    sh.Range("Q32").MergeArea.ClearContents
    sh.Range("Q36").MergeArea.ClearContents
    sh.Range("Q38").MergeArea.ClearContents
    sh.Range("Q41").MergeArea.ClearContents
    sh.Range("Q45").MergeArea.ClearContents
    sh.Range("Q46").MergeArea.ClearContents
    sh.Range("R53").MergeArea.ClearContents
    sh.Range("R54").MergeArea.ClearContents
    sh.Range("R55").MergeArea.ClearContents
    AddCommandButton sh, "C65", "登録・更新", "SaveCurrentCase"
    AddCommandButton sh, "N65", "検索へ戻る", "GoToSearch"
End Sub

Private Sub BuildAddressSettings(ByVal sh As Worksheet)
    sh.Range("A1:C1").Value2 = Array("署（識別名）", "フォルダパス", "ファイル名")
    sh.Range("A1:C1").Font.Bold = True
    sh.Range("A1:C1").Interior.Color = RGB(35, 69, 100)
    sh.Range("A1:C1").Font.Color = vbWhite
    sh.Columns("A").ColumnWidth = 18
    sh.Columns("B").ColumnWidth = 72
    sh.Columns("C").ColumnWidth = 32
    sh.Range("A2:A8").Value2 = Application.Transpose(Array("ippann", "zentai", "Numadu", "Fuji", "Mishima", "Atami", "Shimoda"))
    sh.Range("C3:C8").Value2 = Application.Transpose(Array("進捗整理簿【全体】", "進捗整理簿【沼津】", "進捗整理簿【富士】", "進捗整理簿【三島】", "進捗整理簿【熱海】", "進捗整理簿【下田】"))
    sh.Range("B2:B8").Interior.Color = RGB(255, 250, 210)
    sh.Range("A10").Value2 = "B列に各ファイルがあるフォルダパスを入力。一般事務整理簿のC2は空欄で運用します。"
    sh.Range("A2:C8").NumberFormat = "@"
End Sub

Private Sub BuildFieldSettings(ByVal sh As Worksheet)
    Dim fieldRows(1 To 48, 1 To 4) As Variant
    fieldRows(1, 1) = "番号"
    fieldRows(1, 2) = "A"
    fieldRows(1, 4) = "○"
    fieldRows(2, 1) = "局署番号"
    fieldRows(2, 2) = "B"
    fieldRows(2, 4) = "○"
    fieldRows(3, 1) = "署名"
    fieldRows(3, 2) = "C"
    fieldRows(3, 3) = "F14"
    fieldRows(3, 4) = "○"
    fieldRows(4, 1) = "業務センター"
    fieldRows(4, 2) = "D"
    fieldRows(4, 3) = "F45"
    fieldRows(4, 4) = "○"
    fieldRows(5, 1) = "共通番号枝番号"
    fieldRows(5, 2) = "E"
    fieldRows(5, 3) = "F17"
    fieldRows(5, 4) = "○"
    fieldRows(6, 1) = "共通番号"
    fieldRows(6, 2) = "F"
    fieldRows(6, 4) = "○"
    fieldRows(7, 1) = "枝番号"
    fieldRows(7, 2) = "G"
    fieldRows(7, 4) = "○"
    fieldRows(8, 1) = "税目"
    fieldRows(8, 2) = "H"
    fieldRows(8, 3) = "F27"
    fieldRows(8, 4) = "○"
    fieldRows(9, 1) = "年分・事業年度等"
    fieldRows(9, 2) = "I"
    fieldRows(9, 3) = "Q27"
    fieldRows(9, 4) = "○"
    fieldRows(10, 1) = "申請書種別"
    fieldRows(10, 2) = "J"
    fieldRows(10, 3) = "F29"
    fieldRows(10, 4) = "○"
    fieldRows(11, 1) = "申請日"
    fieldRows(11, 2) = "K"
    fieldRows(11, 3) = "F32"
    fieldRows(11, 4) = "○"
    fieldRows(12, 1) = "提出形式"
    fieldRows(12, 2) = "L"
    fieldRows(12, 3) = "Q29"
    fieldRows(12, 4) = "○"
    fieldRows(13, 1) = "氏名法人名"
    fieldRows(13, 2) = "M"
    fieldRows(13, 3) = "Q14"
    fieldRows(13, 4) = "○"
    fieldRows(14, 1) = "郵便番号"
    fieldRows(14, 2) = "N"
    fieldRows(14, 3) = "F20"
    fieldRows(14, 4) = "○"
    fieldRows(15, 1) = "納税地"
    fieldRows(15, 2) = "O"
    fieldRows(15, 3) = "F23"
    fieldRows(15, 4) = "○"
    fieldRows(16, 1) = "電話番号"
    fieldRows(16, 2) = "P"
    fieldRows(16, 3) = "Q20"
    fieldRows(16, 4) = "○"
    fieldRows(17, 1) = "所管区分"
    fieldRows(17, 2) = "Q"
    fieldRows(17, 4) = "○"
    fieldRows(18, 1) = "調査課部番号"
    fieldRows(18, 2) = "R"
    fieldRows(18, 4) = "○"
    fieldRows(19, 1) = "調査課部門番号"
    fieldRows(19, 2) = "S"
    fieldRows(19, 4) = "○"
    fieldRows(20, 1) = "被合併法人共通番号"
    fieldRows(20, 2) = "T"
    fieldRows(20, 4) = "○"
    fieldRows(21, 1) = "被合併法人枝番号"
    fieldRows(21, 2) = "U"
    fieldRows(21, 4) = "○"
    fieldRows(22, 1) = "申請又は請求の理由"
    fieldRows(22, 2) = "V"
    fieldRows(22, 3) = "Q32"
    fieldRows(22, 4) = "○"
    fieldRows(23, 1) = "処理期限"
    fieldRows(23, 2) = "W"
    fieldRows(23, 3) = "F46"
    fieldRows(23, 4) = "○"
    fieldRows(24, 1) = "処理担当者"
    fieldRows(24, 2) = "X"
    fieldRows(24, 3) = "Q46"
    fieldRows(24, 4) = "○"
    fieldRows(25, 1) = "処理区分"
    fieldRows(25, 2) = "Y"
    fieldRows(25, 3) = "F47"
    fieldRows(25, 4) = "○"
    fieldRows(26, 1) = "処理部署"
    fieldRows(26, 2) = "Z"
    fieldRows(26, 4) = "○"
    fieldRows(27, 1) = "処理グループ（部門）"
    fieldRows(27, 2) = "AA"
    fieldRows(27, 3) = "F48"
    fieldRows(27, 4) = "○"
    fieldRows(28, 1) = "署引継日"
    fieldRows(28, 2) = "AB"
    fieldRows(28, 3) = "F36"
    fieldRows(28, 4) = "○"
    fieldRows(29, 1) = "署引継理由"
    fieldRows(29, 2) = "AC"
    fieldRows(29, 3) = "Q36"
    fieldRows(29, 4) = "○"
    fieldRows(30, 1) = "取下書収受日"
    fieldRows(30, 2) = "AD"
    fieldRows(30, 3) = "F38"
    fieldRows(30, 4) = "○"
    fieldRows(31, 1) = "処理年月日"
    fieldRows(31, 2) = "AE"
    fieldRows(31, 3) = "Q38"
    fieldRows(31, 4) = "○"
    fieldRows(32, 1) = "施行年月日"
    fieldRows(32, 2) = "AF"
    fieldRows(32, 3) = "F41"
    fieldRows(32, 4) = "○"
    fieldRows(33, 1) = "電子通知"
    fieldRows(33, 2) = "AG"
    fieldRows(33, 3) = "Q41"
    fieldRows(33, 4) = "○"
    fieldRows(34, 1) = "調査課連絡用"
    fieldRows(34, 2) = "AH"
    fieldRows(34, 4) = "○"
    fieldRows(35, 1) = "仮装経理フラグ"
    fieldRows(35, 2) = "AI"
    fieldRows(35, 4) = "○"
    fieldRows(36, 1) = "仮装経理に係る最終処理年月日"
    fieldRows(36, 2) = "AJ"
    fieldRows(36, 4) = "○"
    fieldRows(37, 1) = "摘要"
    fieldRows(37, 2) = "AK"
    fieldRows(37, 3) = "F43"
    fieldRows(37, 4) = "○"
    fieldRows(38, 1) = "追加項目"
    fieldRows(38, 2) = "AL"
    fieldRows(38, 4) = "○"
    fieldRows(39, 1) = "追加項目"
    fieldRows(39, 2) = "AM"
    fieldRows(39, 4) = "○"
    fieldRows(40, 1) = "追加項目"
    fieldRows(40, 2) = "AN"
    fieldRows(40, 4) = "○"
    fieldRows(41, 1) = "元ファイル"
    fieldRows(41, 2) = "AO"
    fieldRows(41, 3) = "Q45"
    fieldRows(41, 4) = "○"
    fieldRows(42, 1) = "転入"
    fieldRows(42, 2) = "AP"
    fieldRows(42, 4) = "○"
    fieldRows(43, 1) = "メモ1"
    fieldRows(43, 2) = "AQ"
    fieldRows(43, 3) = "D53"
    fieldRows(43, 4) = "×"
    fieldRows(44, 1) = "メモ2"
    fieldRows(44, 2) = "AR"
    fieldRows(44, 3) = "D54"
    fieldRows(44, 4) = "×"
    fieldRows(45, 1) = "メモ3"
    fieldRows(45, 2) = "AS"
    fieldRows(45, 3) = "D55"
    fieldRows(45, 4) = "×"
    fieldRows(46, 1) = "発送1"
    fieldRows(46, 2) = "AT"
    fieldRows(46, 3) = "R53"
    fieldRows(46, 4) = "×"
    fieldRows(47, 1) = "発送2"
    fieldRows(47, 2) = "AU"
    fieldRows(47, 3) = "R54"
    fieldRows(47, 4) = "×"
    fieldRows(48, 1) = "発送3"
    fieldRows(48, 2) = "AV"
    fieldRows(48, 3) = "R55"
    fieldRows(48, 4) = "×"
    sh.Range("A1:D1").Value2 = Array("新管理簿の項目名（上から列順）", "対応する列", "詳細表示セル", "5署へ転記")
    sh.Range("A2:D49").Value2 = fieldRows
    sh.Range("A1:D1").Font.Bold = True
    sh.Range("A1:D1").Interior.Color = RGB(35, 69, 100)
    sh.Range("A1:D1").Font.Color = vbWhite
    sh.Columns("A").ColumnWidth = 42
    sh.Columns("B").ColumnWidth = 14
    sh.Columns("C").ColumnWidth = 18
    sh.Columns("D").ColumnWidth = 17
    sh.Range("C2:D49").Interior.Color = RGB(255, 250, 210)
    sh.Range("A1:D49").Borders.LineStyle = xlContinuous
    sh.Range("A1:D49").Borders.Color = RGB(221, 226, 232)
    sh.Range("A1:D1").AutoFilter
End Sub

Private Sub AddCommandButton(ByVal sh As Worksheet, ByVal cellAddress As String, _
                             ByVal caption As String, ByVal macroName As String)
    Dim anchor As Range, buttonShape As Shape
    Set anchor = sh.Range(cellAddress)
    Set buttonShape = sh.Shapes.AddShape(msoShapeRoundedRectangle, anchor.Left, anchor.Top, 100, 26)
    buttonShape.TextFrame.Characters.Text = caption
    buttonShape.TextFrame.Characters.Font.Name = "Meiryo UI"
    buttonShape.TextFrame.Characters.Font.Size = 9
    buttonShape.TextFrame.Characters.Font.Color = vbWhite
    buttonShape.Fill.ForeColor.RGB = RGB(35, 69, 100)
    buttonShape.Line.Visible = msoFalse
    buttonShape.OnAction = macroName
End Sub

' 詳細から検索画面へ戻る。
Public Sub GoToSearch()
    ThisWorkbook.Worksheets("検索").Activate
End Sub

' 6つの空の進捗整理簿を生成。Addressに保存先のフォルダパスを記入してから実行する。
Public Sub CreateEmptyLedgers()
    Dim settings As Worksheet, definitions As Worksheet
    Dim fieldRow As Long, settingRow As Long, targetColumn As Long
    Dim folderPath As String, fileName As String, filePath As String
    Dim book As Workbook, sheet As Worksheet, fieldName As String
    Dim isMaster As Boolean, outputName As String
    Set settings = ThisWorkbook.Worksheets("Address")
    Set definitions = ThisWorkbook.Worksheets("Sheet2")
    On Error GoTo Failed
    For settingRow = 3 To 8
        folderPath = Trim$(CStr(settings.Cells(settingRow, "B").Value2))
        fileName = Trim$(CStr(settings.Cells(settingRow, "C").Value2))
        If Len(folderPath) = 0 Or Len(fileName) = 0 Then Err.Raise vbObjectError + 2001, , "Addressの設定不足: " & settingRow & "行"
        If Right$(folderPath, 1) <> "\" And Right$(folderPath, 1) <> "/" Then folderPath = folderPath & IIf(InStr(1, folderPath, "https://", vbTextCompare) = 1, "/", "\")
        If LCase$(Right$(fileName, 5)) <> ".xlsx" Then fileName = fileName & ".xlsx"
        filePath = folderPath & fileName
        If Len(Dir$(filePath)) > 0 Then Err.Raise vbObjectError + 2002, , "既存ファイルがあるため中止: " & filePath
        Set book = Workbooks.Add(xlWBATWorksheet)
        Set sheet = book.Worksheets(1): sheet.Name = "管理簿"
        isMaster = (settingRow = 3)
        targetColumn = 0
        For fieldRow = 2 To 49
            If isMaster Or CStr(definitions.Cells(fieldRow, "D").Value2) = "○" Then
                targetColumn = targetColumn + 1
                fieldName = CStr(definitions.Cells(fieldRow, "A").Value2)
                sheet.Cells(1, targetColumn).Value2 = fieldName
            End If
        Next fieldRow
        sheet.Range(sheet.Cells(1, 1), sheet.Cells(1, targetColumn)).Interior.Color = RGB(35, 69, 100)
        sheet.Range(sheet.Cells(1, 1), sheet.Cells(1, targetColumn)).Font.Color = vbWhite
        sheet.Range(sheet.Cells(1, 1), sheet.Cells(1, targetColumn)).Font.Bold = True
        sheet.Range(sheet.Cells(1, 1), sheet.Cells(1, targetColumn)).ColumnWidth = 19
        sheet.Range("F:G").NumberFormat = "@"
        sheet.Rows(1).RowHeight = 32
        sheet.Rows(1).WrapText = True
        sheet.Range("A1").Select
        ActiveWindow.FreezePanes = True
        book.SaveAs Filename:=filePath, FileFormat:=xlOpenXMLWorkbook
        book.Close SaveChanges:=False
        Set book = Nothing
        Debug.Print "LEDGER_CREATED=" & filePath
    Next settingRow
    Debug.Print "LEDGERS_BUILD=PASS"
    MsgBox "全体1冊と5署5冊の空の管理簿を作成しました。", vbInformation
    Exit Sub
Failed:
    Dim reason As String: reason = Err.Description
    On Error Resume Next
    If Not book Is Nothing Then book.Close SaveChanges:=False
    On Error GoTo 0
    MsgBox "管理簿作成を中断: " & reason, vbExclamation
End Sub

' 元テンプレートの書式パターンを再現する。
Private Sub ApplyTemplateStyle(ByVal target As Range, ByVal styleNumber As Long)
    Select Case styleNumber
        Case 1
            target.Interior.Color = RGB(23,50,76)
            target.Font.Name = "Meiryo UI"
            target.Font.Size = 11
            target.Font.Bold = True
            target.Font.Color = RGB(255,255,255)
            target.VerticalAlignment = xlCenter
        Case 2
            target.Font.Name = "Meiryo UI"
            target.Font.Size = 11
            target.Font.Color = RGB(0,0,0)
            target.VerticalAlignment = xlCenter
        Case 3
            target.Interior.Color = RGB(240,245,250)
            target.Font.Name = "Meiryo UI"
            target.Font.Size = 10
            target.Font.Color = RGB(0,0,0)
            target.VerticalAlignment = xlCenter
        Case 4
            target.Interior.Color = RGB(240,245,250)
            target.Font.Name = "Meiryo UI"
            target.Font.Size = 10
            target.Font.Color = RGB(29,79,122)
            target.VerticalAlignment = xlCenter
        Case 5
            target.Font.Name = "Meiryo UI"
            target.Font.Size = 10
            target.Font.Color = RGB(0,0,0)
            target.VerticalAlignment = xlCenter
        Case 6
            target.Font.Name = "Meiryo UI"
            target.Font.Size = 10
            target.Font.Color = RGB(29,79,122)
            target.VerticalAlignment = xlCenter
        Case 7
            target.Interior.Color = RGB(241,245,249)
            target.Font.Name = "Meiryo UI"
            target.Font.Size = 10
            target.Font.Color = RGB(0,0,0)
            target.VerticalAlignment = xlCenter
        Case 8
            target.Interior.Color = RGB(241,245,249)
            target.Font.Name = "Meiryo UI"
            target.Font.Size = 10
            target.Font.Color = RGB(0,0,0)
            target.HorizontalAlignment = xlCenter
            target.VerticalAlignment = xlCenter
        Case 9
            target.Interior.Color = RGB(255,255,255)
            target.Font.Name = "Meiryo UI"
            target.Font.Size = 10
            target.Font.Color = RGB(0,0,0)
            target.VerticalAlignment = xlCenter
        Case 10
            target.Interior.Color = RGB(255,255,255)
            target.Font.Name = "Meiryo UI"
            target.Font.Size = 10
            target.Font.Color = RGB(0,0,0)
            target.HorizontalAlignment = xlCenter
            target.VerticalAlignment = xlCenter
        Case 11
            target.Interior.Color = RGB(234,245,235)
            target.Font.Name = "Meiryo UI"
            target.Font.Size = 10
            target.Font.Bold = True
            target.Font.Color = RGB(36,122,59)
            target.HorizontalAlignment = xlCenter
            target.VerticalAlignment = xlCenter
        Case 12
            target.Interior.Color = RGB(221,240,223)
            target.Font.Name = "Meiryo UI"
            target.Font.Size = 10
            target.Font.Color = RGB(0,0,0)
            target.VerticalAlignment = xlCenter
        Case 13
            target.Interior.Color = RGB(221,240,223)
            target.Font.Name = "Meiryo UI"
            target.Font.Size = 10
            target.Font.Bold = True
            target.Font.Color = RGB(36,122,59)
            target.HorizontalAlignment = xlCenter
            target.VerticalAlignment = xlCenter
        Case 14
            target.Interior.Color = RGB(255,255,0)
            target.Font.Name = "Meiryo UI"
            target.Font.Size = 10
            target.Font.Color = RGB(0,0,0)
            target.VerticalAlignment = xlCenter
        Case 15
            target.Interior.Color = RGB(242,246,250)
            target.Font.Name = "Meiryo UI"
            target.Font.Size = 10
            target.Font.Color = RGB(29,51,77)
            target.VerticalAlignment = xlCenter
        Case 17
            target.Interior.Color = RGB(242,246,250)
            target.Font.Name = "Meiryo UI"
            target.Font.Size = 10
            target.Font.Color = RGB(23,45,74)
            target.HorizontalAlignment = xlLeft
            target.VerticalAlignment = xlCenter
        Case 18
            target.Interior.Color = RGB(242,246,250)
            target.Font.Name = "Meiryo UI"
            target.Font.Size = 10
            target.Font.Color = RGB(23,45,74)
            target.VerticalAlignment = xlCenter
        Case 19
            target.Interior.Color = RGB(242,246,250)
            target.Font.Name = "Meiryo UI"
            target.Font.Size = 10
            target.Font.Color = RGB(23,45,74)
            target.HorizontalAlignment = xlCenter
            target.VerticalAlignment = xlCenter
        Case 20
            target.Interior.Color = RGB(242,246,250)
            target.Font.Name = "Meiryo UI"
            target.Font.Size = 10
            target.Font.Bold = True
            target.Font.Color = RGB(255,255,255)
            target.HorizontalAlignment = xlLeft
            target.VerticalAlignment = xlCenter
        Case 21
            target.Interior.Color = RGB(242,246,250)
            target.Font.Name = "Meiryo UI"
            target.Font.Size = 10
            target.Font.Bold = True
            target.Font.Color = RGB(255,255,255)
            target.HorizontalAlignment = xlRight
            target.VerticalAlignment = xlCenter
        Case 22
            target.Interior.Color = RGB(242,246,250)
            target.Font.Name = "Meiryo UI"
            target.Font.Size = 10
            target.Font.Color = RGB(23,45,74)
            target.HorizontalAlignment = xlLeft
            target.VerticalAlignment = xlCenter
        Case 23
            target.Interior.Color = RGB(242,246,250)
            target.Font.Name = "Meiryo UI"
            target.Font.Size = 10
            target.Font.Bold = True
            target.Font.Color = RGB(23,45,74)
            target.HorizontalAlignment = xlLeft
            target.VerticalAlignment = xlCenter
            target.WrapText = True
        Case 24
            target.Interior.Color = RGB(255,255,255)
            target.Font.Name = "Meiryo UI"
            target.Font.Size = 10
            target.Font.Bold = True
            target.Font.Color = RGB(23,45,74)
            target.HorizontalAlignment = xlLeft
            target.VerticalAlignment = xlCenter
            target.WrapText = True
        Case 25
            target.Interior.Color = RGB(242,246,250)
            target.Font.Name = "Meiryo UI"
            target.Font.Size = 12
            target.Font.Bold = True
            target.Font.Color = RGB(255,255,255)
            target.HorizontalAlignment = xlLeft
            target.VerticalAlignment = xlCenter
        Case 26
            target.Interior.Color = RGB(242,246,250)
            target.Font.Name = "Meiryo UI"
            target.Font.Size = 9
            target.Font.Color = RGB(255,255,255)
            target.HorizontalAlignment = xlRight
            target.VerticalAlignment = xlCenter
        Case 27
            target.Interior.Color = RGB(242,246,250)
            target.Font.Name = "Meiryo UI"
            target.Font.Size = 11
            target.Font.Color = RGB(23,45,74)
            target.HorizontalAlignment = xlLeft
            target.VerticalAlignment = xlCenter
            target.WrapText = True
        Case 28
            target.Interior.Color = RGB(242,246,250)
            target.Font.Name = "Carlito"
            target.Font.Size = 11
        Case 29
            target.Interior.Color = RGB(239,245,251)
            target.Font.Name = "Meiryo UI"
            target.Font.Size = 11
            target.VerticalAlignment = xlCenter
        Case 30
            target.Interior.Color = RGB(214,229,246)
            target.Font.Name = "Meiryo UI"
            target.Font.Size = 11
            target.Font.Bold = True
            target.Font.Color = RGB(24,48,80)
            target.HorizontalAlignment = xlCenter
            target.VerticalAlignment = xlCenter
            target.Borders(xlEdgeLeft).LineStyle = xlContinuous
            target.Borders(xlEdgeTop).LineStyle = xlContinuous
        Case 31
            target.Interior.Color = RGB(214,229,246)
            target.Font.Name = "Meiryo UI"
            target.Font.Size = 11
            target.Font.Bold = True
            target.Font.Color = RGB(24,48,80)
            target.VerticalAlignment = xlTop
            target.Borders(xlEdgeLeft).LineStyle = xlContinuous
        Case 32
            target.Interior.Color = RGB(214,229,246)
            target.Font.Name = "Meiryo UI"
            target.Font.Size = 11
            target.Font.Bold = True
            target.Font.Color = RGB(24,48,80)
            target.VerticalAlignment = xlTop
        Case 33
            target.Interior.Color = RGB(214,229,246)
            target.Font.Name = "Meiryo UI"
            target.Font.Size = 11
            target.Font.Bold = True
            target.Font.Color = RGB(24,48,80)
            target.VerticalAlignment = xlTop
            target.Borders(xlEdgeRight).LineStyle = xlContinuous
        Case 34
            target.Interior.Color = RGB(239,245,251)
            target.Font.Name = "Meiryo UI"
            target.Font.Size = 11
        Case 35
            target.Interior.Color = RGB(214,229,246)
            target.Font.Name = "Meiryo UI"
            target.Font.Size = 11
            target.Font.Bold = True
            target.Font.Color = RGB(24,48,80)
            target.HorizontalAlignment = xlCenter
            target.VerticalAlignment = xlCenter
            target.Borders(xlEdgeLeft).LineStyle = xlContinuous
        Case 36
            target.Interior.Color = RGB(214,229,246)
            target.Font.Name = "Meiryo UI"
            target.Font.Size = 11
            target.Font.Bold = True
            target.Font.Color = RGB(24,48,80)
            target.HorizontalAlignment = xlCenter
            target.VerticalAlignment = xlCenter
        Case 37
            target.Interior.Color = RGB(214,229,246)
            target.Font.Name = "Meiryo UI"
            target.Font.Size = 11
            target.Font.Bold = True
            target.Font.Color = RGB(24,48,80)
            target.HorizontalAlignment = xlCenter
            target.VerticalAlignment = xlCenter
            target.Borders(xlEdgeRight).LineStyle = xlContinuous
        Case 38
            target.Interior.Color = RGB(239,245,251)
            target.Font.Name = "Meiryo UI"
            target.Font.Size = 11
            target.Borders(xlEdgeLeft).LineStyle = xlContinuous
        Case 39
            target.Interior.Color = RGB(239,245,251)
            target.Font.Name = "Meiryo UI"
            target.Font.Size = 11
            target.Borders(xlEdgeRight).LineStyle = xlContinuous
        Case 40
            target.Interior.Color = RGB(214,229,246)
            target.Font.Name = "Meiryo UI"
            target.Font.Size = 11
            target.Font.Bold = True
            target.Font.Color = RGB(24,48,80)
            target.HorizontalAlignment = xlRight
            target.VerticalAlignment = xlCenter
            target.Borders(xlEdgeLeft).LineStyle = xlContinuous
        Case 41
            target.Interior.Color = RGB(214,229,246)
            target.Font.Name = "Meiryo UI"
            target.Font.Size = 11
            target.Font.Bold = True
            target.Font.Color = RGB(24,48,80)
            target.HorizontalAlignment = xlRight
            target.VerticalAlignment = xlCenter
            target.Borders(xlEdgeLeft).LineStyle = xlContinuous
            target.Borders(xlEdgeBottom).LineStyle = xlContinuous
        Case 42
            target.Interior.Color = RGB(239,245,251)
            target.Font.Name = "Meiryo UI"
            target.Font.Size = 11
            target.Borders(xlEdgeBottom).LineStyle = xlContinuous
        Case 43
            target.Font.Name = "Meiryo UI"
            target.Font.Size = 11
        Case 44
            target.Interior.Color = RGB(249,251,253)
            target.Font.Name = "Carlito"
            target.Font.Size = 11
            target.Borders(xlEdgeLeft).LineStyle = xlContinuous
            target.Borders(xlEdgeTop).LineStyle = xlContinuous
        Case 45
            target.Interior.Color = RGB(249,251,253)
            target.Font.Name = "Carlito"
            target.Font.Size = 11
            target.Borders(xlEdgeTop).LineStyle = xlContinuous
        Case 46
            target.Interior.Color = RGB(249,251,253)
            target.Font.Name = "Carlito"
            target.Font.Size = 11
            target.Borders(xlEdgeRight).LineStyle = xlContinuous
            target.Borders(xlEdgeTop).LineStyle = xlContinuous
        Case 47
            target.Interior.Color = RGB(249,251,253)
            target.Font.Name = "Carlito"
            target.Font.Size = 11
            target.Borders(xlEdgeLeft).LineStyle = xlContinuous
        Case 48
            target.Interior.Color = RGB(249,251,253)
            target.Font.Name = "Carlito"
            target.Font.Size = 11
        Case 49
            target.Interior.Color = RGB(249,251,253)
            target.Font.Name = "Carlito"
            target.Font.Size = 11
            target.Borders(xlEdgeRight).LineStyle = xlContinuous
        Case 50
            target.Interior.Color = RGB(242,246,250)
            target.Font.Name = "Meiryo UI"
            target.Font.Size = 10
            target.Font.Color = RGB(23,45,74)
            target.VerticalAlignment = xlCenter
        Case 51
            target.Interior.Color = RGB(249,251,253)
            target.Font.Name = "Carlito"
            target.Font.Size = 11
            target.Borders(xlEdgeLeft).LineStyle = xlContinuous
            target.Borders(xlEdgeBottom).LineStyle = xlContinuous
        Case 52
            target.Interior.Color = RGB(249,251,253)
            target.Font.Name = "Carlito"
            target.Font.Size = 11
            target.Borders(xlEdgeBottom).LineStyle = xlContinuous
        Case 53
            target.Interior.Color = RGB(249,251,253)
            target.Font.Name = "Carlito"
            target.Font.Size = 11
            target.Borders(xlEdgeRight).LineStyle = xlContinuous
            target.Borders(xlEdgeBottom).LineStyle = xlContinuous
        Case 54
            target.Font.Name = "Carlito"
            target.Font.Size = 11
        Case 55
            target.Interior.Color = RGB(243,246,249)
            target.Font.Name = "Meiryo UI"
            target.Font.Size = 10
        Case 56
            target.Interior.Color = RGB(214,239,214)
            target.Font.Name = "Meiryo UI"
            target.Font.Size = 10
        Case 57
            target.Interior.Color = RGB(255,255,255)
            target.Font.Name = "Meiryo UI"
            target.Font.Size = 10
        Case 58
            target.Interior.Color = RGB(234,240,246)
            target.Font.Name = "Meiryo UI"
            target.Font.Size = 10
        Case 59
            target.Interior.Color = RGB(242,246,250)
            target.Font.Name = "Meiryo UI"
            target.Font.Size = 9
            target.Font.Bold = True
            target.Font.Color = RGB(23,45,74)
            target.VerticalAlignment = xlCenter
        Case 60
            target.Interior.Color = RGB(31,73,125)
            target.Font.Name = "Meiryo UI"
            target.Font.Size = 10
            target.Font.Color = RGB(0,0,0)
            target.VerticalAlignment = xlCenter
        Case 61
            target.Interior.Color = RGB(233,239,246)
            target.Font.Name = "Meiryo UI"
            target.Font.Size = 10
            target.Font.Color = RGB(23,45,74)
            target.HorizontalAlignment = xlLeft
            target.VerticalAlignment = xlCenter
        Case 62
            target.Interior.Color = RGB(242,246,250)
            target.Font.Name = "Meiryo UI"
            target.Font.Size = 10
            target.Font.Color = RGB(23,45,74)
            target.HorizontalAlignment = xlLeft
            target.VerticalAlignment = xlCenter
        Case 63
            target.Interior.Color = RGB(23,45,74)
            target.Font.Name = "Meiryo UI"
            target.Font.Size = 19
            target.Font.Bold = True
            target.Font.Color = RGB(255,255,255)
            target.HorizontalAlignment = xlLeft
            target.VerticalAlignment = xlCenter
        Case 64
            target.Interior.Color = RGB(242,246,250)
            target.Font.Name = "Meiryo UI"
            target.Font.Size = 13
            target.Font.Bold = True
            target.Font.Color = RGB(23,45,74)
            target.HorizontalAlignment = xlLeft
            target.VerticalAlignment = xlCenter
        Case 65
            target.Interior.Color = RGB(255,255,255)
            target.Font.Name = "Meiryo UI"
            target.Font.Size = 11
            target.Font.Color = RGB(23,45,74)
            target.HorizontalAlignment = xlLeft
            target.VerticalAlignment = xlCenter
        Case 66
            target.Interior.Color = RGB(223,236,247)
            target.Font.Name = "Meiryo UI"
            target.Font.Size = 11
            target.Font.Bold = True
            target.Font.Color = RGB(23,45,74)
            target.HorizontalAlignment = xlLeft
            target.VerticalAlignment = xlCenter
        Case 67
            target.Interior.Color = RGB(23,45,74)
            target.Font.Name = "Meiryo UI"
            target.Font.Size = 10
            target.Font.Bold = True
            target.Font.Color = RGB(255,255,255)
            target.HorizontalAlignment = xlCenter
            target.VerticalAlignment = xlCenter
        Case 68
            target.Interior.Color = RGB(255,255,255)
            target.Font.Name = "Meiryo UI"
            target.Font.Size = 10
            target.Font.Color = RGB(23,45,74)
            target.HorizontalAlignment = xlLeft
            target.VerticalAlignment = xlCenter
        Case 69
            target.Interior.Color = RGB(234,240,246)
            target.Font.Name = "Meiryo UI"
            target.Font.Size = 10
            target.Font.Color = RGB(23,45,74)
            target.HorizontalAlignment = xlLeft
            target.VerticalAlignment = xlCenter
        Case 70
            target.Interior.Color = RGB(223,236,247)
            target.Font.Name = "Meiryo UI"
            target.Font.Size = 10
            target.Font.Bold = True
            target.Font.Color = RGB(23,45,74)
            target.HorizontalAlignment = xlLeft
            target.VerticalAlignment = xlCenter
        Case 71
            target.Interior.Color = RGB(223,236,247)
            target.Font.Name = "Meiryo UI"
            target.Font.Size = 10
            target.Font.Bold = True
            target.Font.Color = RGB(23,45,74)
            target.HorizontalAlignment = xlLeft
            target.VerticalAlignment = xlCenter
        Case 72
            target.Interior.Color = RGB(23,45,74)
            target.Font.Name = "Meiryo UI"
            target.Font.Size = 12
            target.Font.Bold = True
            target.Font.Color = RGB(255,255,255)
            target.HorizontalAlignment = xlLeft
            target.VerticalAlignment = xlCenter
        Case 73
            target.Interior.Color = RGB(23,45,74)
            target.Font.Name = "Meiryo UI"
            target.Font.Size = 9
            target.Font.Color = RGB(255,255,255)
            target.HorizontalAlignment = xlRight
            target.VerticalAlignment = xlCenter
        Case 74
            target.Interior.Color = RGB(218,233,247)
            target.Font.Name = "Meiryo UI"
            target.Font.Size = 11
            target.Font.Bold = True
            target.Font.Color = RGB(16,39,68)
            target.HorizontalAlignment = xlLeft
            target.VerticalAlignment = xlCenter
            target.WrapText = True
        Case 75
            target.Interior.Color = RGB(218,233,247)
            target.Font.Name = "Meiryo UI"
            target.Font.Size = 11
            target.Font.Bold = True
            target.Font.Color = RGB(16,39,68)
            target.HorizontalAlignment = xlLeft
            target.VerticalAlignment = xlCenter
            target.WrapText = True
        Case 76
            target.Interior.Color = RGB(218,233,247)
            target.Font.Name = "Meiryo UI"
            target.Font.Size = 11
            target.Font.Bold = True
            target.Font.Color = RGB(16,39,68)
            target.HorizontalAlignment = xlLeft
            target.VerticalAlignment = xlCenter
        Case 77
            target.Interior.Color = RGB(218,233,247)
            target.Font.Name = "Meiryo UI"
            target.Font.Size = 12
            target.Font.Bold = True
            target.Font.Color = RGB(16,39,68)
            target.HorizontalAlignment = xlLeft
            target.VerticalAlignment = xlCenter
        Case 78
            target.Interior.Color = RGB(23,45,74)
            target.Font.Name = "Meiryo UI"
            target.Font.Size = 9
            target.Font.Bold = True
            target.Font.Color = RGB(255,255,255)
            target.HorizontalAlignment = xlLeft
            target.VerticalAlignment = xlCenter
        Case 79
            target.Interior.Color = RGB(230,246,220)
            target.Font.Name = "Meiryo UI"
            target.Font.Size = 11
            target.Font.Bold = True
            target.Font.Color = RGB(16,39,68)
            target.HorizontalAlignment = xlLeft
            target.VerticalAlignment = xlCenter
            target.WrapText = True
            target.Borders(xlEdgeLeft).LineStyle = xlContinuous
            target.Borders(xlEdgeTop).LineStyle = xlContinuous
        Case 80
            target.Interior.Color = RGB(230,246,220)
            target.Font.Name = "Meiryo UI"
            target.Font.Size = 11
            target.Font.Bold = True
            target.Font.Color = RGB(16,39,68)
            target.HorizontalAlignment = xlLeft
            target.VerticalAlignment = xlCenter
            target.WrapText = True
            target.Borders(xlEdgeTop).LineStyle = xlContinuous
        Case 81
            target.Interior.Color = RGB(230,246,220)
            target.Font.Name = "Meiryo UI"
            target.Font.Size = 11
            target.Font.Bold = True
            target.Font.Color = RGB(16,39,68)
            target.HorizontalAlignment = xlLeft
            target.VerticalAlignment = xlCenter
            target.WrapText = True
            target.Borders(xlEdgeRight).LineStyle = xlContinuous
            target.Borders(xlEdgeTop).LineStyle = xlContinuous
        Case 82
            target.Interior.Color = RGB(230,246,220)
            target.Font.Name = "Meiryo UI"
            target.Font.Size = 11
            target.Font.Bold = True
            target.Font.Color = RGB(16,39,68)
            target.HorizontalAlignment = xlLeft
            target.VerticalAlignment = xlCenter
            target.WrapText = True
            target.Borders(xlEdgeLeft).LineStyle = xlContinuous
            target.Borders(xlEdgeBottom).LineStyle = xlContinuous
        Case 83
            target.Interior.Color = RGB(230,246,220)
            target.Font.Name = "Meiryo UI"
            target.Font.Size = 11
            target.Font.Bold = True
            target.Font.Color = RGB(16,39,68)
            target.HorizontalAlignment = xlLeft
            target.VerticalAlignment = xlCenter
            target.WrapText = True
            target.Borders(xlEdgeBottom).LineStyle = xlContinuous
        Case 84
            target.Interior.Color = RGB(230,246,220)
            target.Font.Name = "Meiryo UI"
            target.Font.Size = 11
            target.Font.Bold = True
            target.Font.Color = RGB(16,39,68)
            target.HorizontalAlignment = xlLeft
            target.VerticalAlignment = xlCenter
            target.WrapText = True
            target.Borders(xlEdgeRight).LineStyle = xlContinuous
            target.Borders(xlEdgeBottom).LineStyle = xlContinuous
        Case 85
            target.Interior.Color = RGB(218,233,247)
            target.Font.Name = "Meiryo UI"
            target.Font.Size = 10
            target.Font.Bold = True
            target.Font.Color = RGB(16,39,68)
            target.HorizontalAlignment = xlLeft
            target.VerticalAlignment = xlCenter
        Case 86
            target.Interior.Color = RGB(255,255,255)
            target.Font.Name = "Meiryo UI"
            target.Font.Size = 11
            target.Font.Color = RGB(23,45,74)
            target.HorizontalAlignment = xlLeft
            target.VerticalAlignment = xlCenter
            target.Borders(xlEdgeLeft).LineStyle = xlContinuous
            target.Borders(xlEdgeTop).LineStyle = xlContinuous
            target.Borders(xlEdgeBottom).LineStyle = xlContinuous
        Case 87
            target.Interior.Color = RGB(255,255,255)
            target.Font.Name = "Meiryo UI"
            target.Font.Size = 11
            target.Font.Color = RGB(23,45,74)
            target.HorizontalAlignment = xlLeft
            target.VerticalAlignment = xlCenter
            target.WrapText = True
            target.Borders(xlEdgeTop).LineStyle = xlContinuous
        Case 88
            target.Interior.Color = RGB(255,255,255)
            target.Font.Name = "Meiryo UI"
            target.Font.Size = 11
            target.Font.Color = RGB(23,45,74)
            target.HorizontalAlignment = xlLeft
            target.VerticalAlignment = xlCenter
            target.WrapText = True
            target.Borders(xlEdgeRight).LineStyle = xlContinuous
            target.Borders(xlEdgeTop).LineStyle = xlContinuous
        Case 89
            target.Interior.Color = RGB(255,255,255)
            target.Font.Name = "Meiryo UI"
            target.Font.Size = 11
            target.Font.Color = RGB(23,45,74)
            target.HorizontalAlignment = xlLeft
            target.VerticalAlignment = xlCenter
            target.WrapText = True
            target.Borders(xlEdgeLeft).LineStyle = xlContinuous
            target.Borders(xlEdgeBottom).LineStyle = xlContinuous
        Case 90
            target.Interior.Color = RGB(255,255,255)
            target.Font.Name = "Meiryo UI"
            target.Font.Size = 11
            target.Font.Color = RGB(23,45,74)
            target.HorizontalAlignment = xlLeft
            target.VerticalAlignment = xlCenter
            target.WrapText = True
            target.Borders(xlEdgeBottom).LineStyle = xlContinuous
        Case 91
            target.Interior.Color = RGB(255,255,255)
            target.Font.Name = "Meiryo UI"
            target.Font.Size = 11
            target.Font.Color = RGB(23,45,74)
            target.HorizontalAlignment = xlLeft
            target.VerticalAlignment = xlCenter
            target.WrapText = True
            target.Borders(xlEdgeRight).LineStyle = xlContinuous
            target.Borders(xlEdgeBottom).LineStyle = xlContinuous
        Case 92
            target.Interior.Color = RGB(255,255,255)
            target.Font.Name = "Meiryo UI"
            target.Font.Size = 11
            target.Font.Color = RGB(23,45,74)
            target.HorizontalAlignment = xlLeft
            target.VerticalAlignment = xlCenter
            target.Borders(xlEdgeTop).LineStyle = xlContinuous
        Case 93
            target.Interior.Color = RGB(255,255,255)
            target.Font.Name = "Meiryo UI"
            target.Font.Size = 11
            target.Font.Color = RGB(23,45,74)
            target.HorizontalAlignment = xlLeft
            target.VerticalAlignment = xlCenter
            target.Borders(xlEdgeRight).LineStyle = xlContinuous
            target.Borders(xlEdgeTop).LineStyle = xlContinuous
        Case 94
            target.Interior.Color = RGB(255,255,255)
            target.Font.Name = "Meiryo UI"
            target.Font.Size = 11
            target.Font.Color = RGB(23,45,74)
            target.HorizontalAlignment = xlLeft
            target.VerticalAlignment = xlCenter
            target.Borders(xlEdgeLeft).LineStyle = xlContinuous
            target.Borders(xlEdgeBottom).LineStyle = xlContinuous
        Case 95
            target.Interior.Color = RGB(255,255,255)
            target.Font.Name = "Meiryo UI"
            target.Font.Size = 11
            target.Font.Color = RGB(23,45,74)
            target.HorizontalAlignment = xlLeft
            target.VerticalAlignment = xlCenter
            target.Borders(xlEdgeBottom).LineStyle = xlContinuous
        Case 96
            target.Interior.Color = RGB(255,255,255)
            target.Font.Name = "Meiryo UI"
            target.Font.Size = 11
            target.Font.Color = RGB(23,45,74)
            target.HorizontalAlignment = xlLeft
            target.VerticalAlignment = xlCenter
            target.Borders(xlEdgeRight).LineStyle = xlContinuous
            target.Borders(xlEdgeBottom).LineStyle = xlContinuous
        Case 97
            target.Interior.Color = RGB(255,255,255)
            target.Font.Name = "Meiryo UI"
            target.Font.Size = 11
            target.Font.Color = RGB(23,45,74)
            target.HorizontalAlignment = xlLeft
            target.VerticalAlignment = xlCenter
            target.WrapText = True
            target.Borders(xlEdgeTop).LineStyle = xlContinuous
        Case 98
            target.Interior.Color = RGB(255,255,255)
            target.Font.Name = "Meiryo UI"
            target.Font.Size = 11
            target.Font.Color = RGB(23,45,74)
            target.HorizontalAlignment = xlLeft
            target.VerticalAlignment = xlCenter
            target.WrapText = True
            target.Borders(xlEdgeRight).LineStyle = xlContinuous
            target.Borders(xlEdgeTop).LineStyle = xlContinuous
        Case 99
            target.Interior.Color = RGB(255,255,255)
            target.Font.Name = "Meiryo UI"
            target.Font.Size = 11
            target.Font.Color = RGB(23,45,74)
            target.HorizontalAlignment = xlLeft
            target.VerticalAlignment = xlCenter
            target.WrapText = True
            target.Borders(xlEdgeLeft).LineStyle = xlContinuous
            target.Borders(xlEdgeBottom).LineStyle = xlContinuous
        Case 100
            target.Interior.Color = RGB(255,255,255)
            target.Font.Name = "Meiryo UI"
            target.Font.Size = 11
            target.Font.Color = RGB(23,45,74)
            target.HorizontalAlignment = xlLeft
            target.VerticalAlignment = xlCenter
            target.WrapText = True
            target.Borders(xlEdgeBottom).LineStyle = xlContinuous
        Case 101
            target.Interior.Color = RGB(255,255,255)
            target.Font.Name = "Meiryo UI"
            target.Font.Size = 11
            target.Font.Color = RGB(23,45,74)
            target.HorizontalAlignment = xlLeft
            target.VerticalAlignment = xlCenter
            target.WrapText = True
            target.Borders(xlEdgeRight).LineStyle = xlContinuous
            target.Borders(xlEdgeBottom).LineStyle = xlContinuous
        Case 102
            target.Interior.Color = RGB(255,255,255)
            target.Font.Name = "Meiryo UI"
            target.Font.Size = 11
            target.Font.Color = RGB(16,39,68)
            target.HorizontalAlignment = xlLeft
            target.VerticalAlignment = xlCenter
            target.WrapText = True
            target.Borders(xlEdgeTop).LineStyle = xlContinuous
        Case 103
            target.Interior.Color = RGB(255,255,255)
            target.Font.Name = "Meiryo UI"
            target.Font.Size = 11
            target.Font.Color = RGB(16,39,68)
            target.HorizontalAlignment = xlLeft
            target.VerticalAlignment = xlCenter
            target.WrapText = True
            target.Borders(xlEdgeRight).LineStyle = xlContinuous
            target.Borders(xlEdgeTop).LineStyle = xlContinuous
        Case 104
            target.Interior.Color = RGB(255,255,255)
            target.Font.Name = "Meiryo UI"
            target.Font.Size = 11
            target.Font.Color = RGB(16,39,68)
            target.HorizontalAlignment = xlLeft
            target.VerticalAlignment = xlCenter
            target.WrapText = True
            target.Borders(xlEdgeLeft).LineStyle = xlContinuous
            target.Borders(xlEdgeBottom).LineStyle = xlContinuous
        Case 105
            target.Interior.Color = RGB(255,255,255)
            target.Font.Name = "Meiryo UI"
            target.Font.Size = 11
            target.Font.Color = RGB(16,39,68)
            target.HorizontalAlignment = xlLeft
            target.VerticalAlignment = xlCenter
            target.WrapText = True
            target.Borders(xlEdgeBottom).LineStyle = xlContinuous
        Case 106
            target.Interior.Color = RGB(255,255,255)
            target.Font.Name = "Meiryo UI"
            target.Font.Size = 11
            target.Font.Color = RGB(16,39,68)
            target.HorizontalAlignment = xlLeft
            target.VerticalAlignment = xlCenter
            target.WrapText = True
            target.Borders(xlEdgeRight).LineStyle = xlContinuous
            target.Borders(xlEdgeBottom).LineStyle = xlContinuous
        Case 107
            target.Interior.Color = RGB(218,233,247)
            target.Font.Name = "Meiryo UI"
            target.Font.Size = 10
            target.Font.Bold = True
            target.Font.Color = RGB(16,39,68)
            target.HorizontalAlignment = xlLeft
            target.VerticalAlignment = xlCenter
            target.WrapText = True
        Case 108
            target.Interior.Color = RGB(255,255,255)
            target.Font.Name = "Meiryo UI"
            target.Font.Size = 11
            target.Font.Color = RGB(23,45,74)
            target.HorizontalAlignment = xlLeft
            target.VerticalAlignment = xlCenter
        Case 109
            target.Interior.Color = RGB(255,255,255)
            target.Font.Name = "Meiryo UI"
            target.Font.Size = 11
            target.Font.Color = RGB(23,45,74)
            target.HorizontalAlignment = xlLeft
            target.VerticalAlignment = xlCenter
        Case 110
            target.Font.Name = "游ゴシック"
            target.Font.Size = 11
            target.Font.Color = RGB(0,0,0)
            target.HorizontalAlignment = xlLeft
            target.VerticalAlignment = xlCenter
        Case 111
            target.Interior.Color = RGB(255,255,255)
            target.Font.Name = "Meiryo UI"
            target.Font.Size = 11
            target.Font.Color = RGB(23,45,74)
            target.HorizontalAlignment = xlLeft
            target.VerticalAlignment = xlCenter
            target.Borders(xlEdgeTop).LineStyle = xlContinuous
            target.Borders(xlEdgeBottom).LineStyle = xlContinuous
        Case 112
            target.Interior.Color = RGB(255,255,255)
            target.Font.Name = "Meiryo UI"
            target.Font.Size = 11
            target.Font.Color = RGB(23,45,74)
            target.HorizontalAlignment = xlLeft
            target.VerticalAlignment = xlCenter
            target.Borders(xlEdgeRight).LineStyle = xlContinuous
            target.Borders(xlEdgeTop).LineStyle = xlContinuous
            target.Borders(xlEdgeBottom).LineStyle = xlContinuous
        Case 113
            target.Interior.Color = RGB(255,255,255)
            target.Font.Name = "Meiryo UI"
            target.Font.Size = 11
            target.Font.Color = RGB(23,45,74)
            target.HorizontalAlignment = xlLeft
            target.VerticalAlignment = xlCenter
            target.Borders(xlEdgeTop).LineStyle = xlContinuous
        Case 114
            target.Interior.Color = RGB(255,255,255)
            target.Font.Name = "Meiryo UI"
            target.Font.Size = 11
            target.Font.Color = RGB(23,45,74)
            target.HorizontalAlignment = xlLeft
            target.VerticalAlignment = xlCenter
            target.Borders(xlEdgeRight).LineStyle = xlContinuous
            target.Borders(xlEdgeTop).LineStyle = xlContinuous
        Case 115
            target.Interior.Color = RGB(255,255,255)
            target.Font.Name = "Meiryo UI"
            target.Font.Size = 11
            target.Font.Color = RGB(23,45,74)
            target.HorizontalAlignment = xlLeft
            target.VerticalAlignment = xlCenter
            target.Borders(xlEdgeLeft).LineStyle = xlContinuous
            target.Borders(xlEdgeBottom).LineStyle = xlContinuous
        Case 116
            target.Interior.Color = RGB(255,255,255)
            target.Font.Name = "Meiryo UI"
            target.Font.Size = 11
            target.Font.Color = RGB(23,45,74)
            target.HorizontalAlignment = xlLeft
            target.VerticalAlignment = xlCenter
            target.Borders(xlEdgeBottom).LineStyle = xlContinuous
        Case 117
            target.Interior.Color = RGB(255,255,255)
            target.Font.Name = "Meiryo UI"
            target.Font.Size = 11
            target.Font.Color = RGB(23,45,74)
            target.HorizontalAlignment = xlLeft
            target.VerticalAlignment = xlCenter
            target.Borders(xlEdgeRight).LineStyle = xlContinuous
            target.Borders(xlEdgeBottom).LineStyle = xlContinuous
        Case 118
            target.Interior.Color = RGB(242,246,250)
            target.Font.Name = "Meiryo UI"
            target.Font.Size = 11
            target.Font.Color = RGB(23,45,74)
            target.HorizontalAlignment = xlLeft
            target.VerticalAlignment = xlCenter
            target.WrapText = True
        Case 119
            target.Interior.Color = RGB(242,246,250)
            target.Font.Name = "Meiryo UI"
            target.Font.Size = 9
            target.Font.Bold = True
            target.Font.Color = RGB(16,39,68)
            target.HorizontalAlignment = xlCenter
            target.VerticalAlignment = xlCenter
            target.WrapText = True
        Case 120
            target.Interior.Color = RGB(255,255,255)
            target.Font.Name = "Meiryo UI"
            target.Font.Size = 11
            target.Font.Color = RGB(23,45,74)
            target.HorizontalAlignment = xlLeft
            target.VerticalAlignment = xlCenter
            target.Borders(xlEdgeTop).LineStyle = xlContinuous
            target.Borders(xlEdgeBottom).LineStyle = xlContinuous
        Case 121
            target.Interior.Color = RGB(255,255,255)
            target.Font.Name = "Meiryo UI"
            target.Font.Size = 11
            target.Font.Color = RGB(23,45,74)
            target.HorizontalAlignment = xlLeft
            target.VerticalAlignment = xlCenter
            target.Borders(xlEdgeRight).LineStyle = xlContinuous
            target.Borders(xlEdgeTop).LineStyle = xlContinuous
            target.Borders(xlEdgeBottom).LineStyle = xlContinuous
        Case 122
            target.Interior.Color = RGB(255,255,255)
            target.Font.Name = "Meiryo UI"
            target.Font.Size = 11
            target.Font.Color = RGB(23,45,74)
            target.HorizontalAlignment = xlLeft
            target.VerticalAlignment = xlCenter
            target.Borders(xlEdgeTop).LineStyle = xlContinuous
            target.Borders(xlEdgeBottom).LineStyle = xlContinuous
        Case 123
            target.Interior.Color = RGB(255,255,255)
            target.Font.Name = "Meiryo UI"
            target.Font.Size = 11
            target.Font.Color = RGB(23,45,74)
            target.HorizontalAlignment = xlLeft
            target.VerticalAlignment = xlCenter
            target.Borders(xlEdgeTop).LineStyle = xlContinuous
        Case 124
            target.Interior.Color = RGB(255,255,255)
            target.Font.Name = "Meiryo UI"
            target.Font.Size = 11
            target.Font.Color = RGB(23,45,74)
            target.HorizontalAlignment = xlLeft
            target.VerticalAlignment = xlCenter
            target.Borders(xlEdgeRight).LineStyle = xlContinuous
            target.Borders(xlEdgeTop).LineStyle = xlContinuous
        Case 125
            target.Interior.Color = RGB(255,255,255)
            target.Font.Name = "Meiryo UI"
            target.Font.Size = 11
            target.Font.Color = RGB(23,45,74)
            target.HorizontalAlignment = xlLeft
            target.VerticalAlignment = xlCenter
            target.Borders(xlEdgeLeft).LineStyle = xlContinuous
            target.Borders(xlEdgeBottom).LineStyle = xlContinuous
        Case 126
            target.Interior.Color = RGB(255,255,255)
            target.Font.Name = "Meiryo UI"
            target.Font.Size = 11
            target.Font.Color = RGB(23,45,74)
            target.HorizontalAlignment = xlLeft
            target.VerticalAlignment = xlCenter
            target.Borders(xlEdgeBottom).LineStyle = xlContinuous
        Case 127
            target.Interior.Color = RGB(255,255,255)
            target.Font.Name = "Meiryo UI"
            target.Font.Size = 11
            target.Font.Color = RGB(23,45,74)
            target.HorizontalAlignment = xlLeft
            target.VerticalAlignment = xlCenter
            target.Borders(xlEdgeRight).LineStyle = xlContinuous
            target.Borders(xlEdgeBottom).LineStyle = xlContinuous
        Case 128
            target.Interior.Color = RGB(255,255,255)
            target.Font.Name = "Meiryo UI"
            target.Font.Size = 11
            target.Font.Color = RGB(23,45,74)
            target.HorizontalAlignment = xlLeft
            target.VerticalAlignment = xlCenter
            target.Borders(xlEdgeLeft).LineStyle = xlContinuous
            target.Borders(xlEdgeRight).LineStyle = xlContinuous
            target.Borders(xlEdgeTop).LineStyle = xlContinuous
            target.Borders(xlEdgeBottom).LineStyle = xlContinuous
        Case 129
            target.Interior.Color = RGB(242,246,250)
            target.Font.Name = "Carlito"
            target.Font.Size = 11
        Case 130
            target.Interior.Color = RGB(218,233,247)
            target.Font.Name = "Meiryo UI"
            target.Font.Size = 10
            target.Font.Bold = True
            target.Font.Color = RGB(16,39,68)
            target.HorizontalAlignment = xlLeft
            target.VerticalAlignment = xlCenter
        Case 131
            target.Interior.Color = RGB(255,255,255)
            target.Font.Name = "Meiryo UI"
            target.Font.Size = 11
            target.Font.Color = RGB(23,45,74)
            target.HorizontalAlignment = xlLeft
            target.VerticalAlignment = xlCenter
            target.Borders(xlEdgeLeft).LineStyle = xlContinuous
            target.Borders(xlEdgeRight).LineStyle = xlContinuous
            target.Borders(xlEdgeTop).LineStyle = xlContinuous
            target.Borders(xlEdgeBottom).LineStyle = xlContinuous
        Case 132
            target.Interior.Color = RGB(255,255,255)
            target.Font.Name = "Carlito"
            target.Font.Size = 11
        Case 133
            target.Interior.Color = RGB(255,255,255)
            target.Font.Name = "Meiryo UI"
            target.Font.Size = 11
            target.Font.Color = RGB(23,45,74)
            target.HorizontalAlignment = xlLeft
            target.VerticalAlignment = xlCenter
            target.Borders(xlEdgeLeft).LineStyle = xlContinuous
        Case 134
            target.Interior.Color = RGB(242,246,250)
            target.Font.Name = "Meiryo UI"
            target.Font.Size = 10
            target.Font.Color = RGB(23,45,74)
            target.VerticalAlignment = xlCenter
        Case 135
            target.Font.Name = "游ゴシック"
            target.Font.Size = 11
            target.Font.Color = RGB(0,0,0)
            target.VerticalAlignment = xlCenter
        Case 136
            target.Font.Name = "Carlito"
            target.Font.Size = 11
        Case 137
            target.Interior.Color = RGB(214,229,246)
            target.Font.Name = "Meiryo UI"
            target.Font.Size = 11
            target.Font.Bold = True
            target.Font.Color = RGB(24,48,80)
            target.VerticalAlignment = xlCenter
            target.Borders(xlEdgeLeft).LineStyle = xlContinuous
            target.Borders(xlEdgeTop).LineStyle = xlContinuous
        Case 138
            target.Interior.Color = RGB(214,229,246)
            target.Font.Name = "Meiryo UI"
            target.Font.Size = 11
            target.Font.Bold = True
            target.Font.Color = RGB(24,48,80)
            target.VerticalAlignment = xlCenter
            target.Borders(xlEdgeTop).LineStyle = xlContinuous
        Case 139
            target.Interior.Color = RGB(214,229,246)
            target.Font.Name = "Meiryo UI"
            target.Font.Size = 11
            target.Font.Bold = True
            target.Font.Color = RGB(24,48,80)
            target.VerticalAlignment = xlCenter
            target.Borders(xlEdgeRight).LineStyle = xlContinuous
            target.Borders(xlEdgeTop).LineStyle = xlContinuous
        Case 140
            target.Interior.Color = RGB(214,229,246)
            target.Font.Name = "Meiryo UI"
            target.Font.Size = 10
            target.Font.Bold = True
            target.Font.Color = RGB(24,48,80)
            target.HorizontalAlignment = xlLeft
            target.VerticalAlignment = xlCenter
            target.Borders(xlEdgeTop).LineStyle = xlContinuous
        Case 141
            target.Interior.Color = RGB(214,229,246)
            target.Font.Name = "Meiryo UI"
            target.Font.Size = 11
            target.Font.Bold = True
            target.Font.Color = RGB(24,48,80)
            target.HorizontalAlignment = xlLeft
            target.VerticalAlignment = xlCenter
            target.Borders(xlEdgeTop).LineStyle = xlContinuous
        Case 142
            target.Interior.Color = RGB(214,229,246)
            target.Font.Name = "Meiryo UI"
            target.Font.Size = 11
            target.Font.Bold = True
            target.Font.Color = RGB(24,48,80)
            target.HorizontalAlignment = xlCenter
            target.VerticalAlignment = xlCenter
            target.Borders(xlEdgeTop).LineStyle = xlContinuous
        Case 143
            target.Interior.Color = RGB(214,229,246)
            target.Font.Name = "Meiryo UI"
            target.Font.Size = 11
            target.Font.Bold = True
            target.Font.Color = RGB(24,48,80)
            target.HorizontalAlignment = xlCenter
            target.VerticalAlignment = xlCenter
            target.Borders(xlEdgeRight).LineStyle = xlContinuous
            target.Borders(xlEdgeTop).LineStyle = xlContinuous
        Case 144
            target.Interior.Color = RGB(255,255,255)
            target.Font.Name = "Meiryo UI"
            target.Font.Size = 11
            target.Font.Color = RGB(23,45,74)
            target.HorizontalAlignment = xlLeft
            target.VerticalAlignment = xlCenter
            target.Borders(xlEdgeLeft).LineStyle = xlContinuous
            target.Borders(xlEdgeBottom).LineStyle = xlContinuous
        Case 145
            target.Font.Name = "游ゴシック"
            target.Font.Size = 11
            target.Font.Color = RGB(0,0,0)
            target.VerticalAlignment = xlCenter
            target.Borders(xlEdgeBottom).LineStyle = xlContinuous
        Case 146
            target.Interior.Color = RGB(255,255,255)
            target.Font.Name = "Meiryo UI"
            target.Font.Size = 11
            target.Font.Color = RGB(23,45,74)
            target.HorizontalAlignment = xlLeft
            target.VerticalAlignment = xlCenter
            target.WrapText = True
            target.Borders(xlEdgeLeft).LineStyle = xlContinuous
            target.Borders(xlEdgeTop).LineStyle = xlContinuous
        Case 147
            target.Interior.Color = RGB(255,255,255)
            target.Font.Name = "Meiryo UI"
            target.Font.Size = 11
            target.Font.Color = RGB(23,45,74)
            target.HorizontalAlignment = xlLeft
            target.VerticalAlignment = xlCenter
            target.WrapText = True
            target.Borders(xlEdgeTop).LineStyle = xlContinuous
        Case 148
            target.Interior.Color = RGB(255,255,255)
            target.Font.Name = "Meiryo UI"
            target.Font.Size = 11
            target.Font.Color = RGB(23,45,74)
            target.HorizontalAlignment = xlLeft
            target.VerticalAlignment = xlCenter
            target.WrapText = True
            target.Borders(xlEdgeRight).LineStyle = xlContinuous
            target.Borders(xlEdgeTop).LineStyle = xlContinuous
        Case 149
            target.Interior.Color = RGB(255,255,255)
            target.Font.Name = "Meiryo UI"
            target.Font.Size = 11
            target.Font.Color = RGB(23,45,74)
            target.HorizontalAlignment = xlLeft
            target.VerticalAlignment = xlCenter
            target.WrapText = True
            target.Borders(xlEdgeLeft).LineStyle = xlContinuous
        Case 150
            target.Interior.Color = RGB(255,255,255)
            target.Font.Name = "Meiryo UI"
            target.Font.Size = 11
            target.Font.Color = RGB(23,45,74)
            target.HorizontalAlignment = xlLeft
            target.VerticalAlignment = xlCenter
            target.WrapText = True
        Case 151
            target.Interior.Color = RGB(255,255,255)
            target.Font.Name = "Meiryo UI"
            target.Font.Size = 11
            target.Font.Color = RGB(23,45,74)
            target.HorizontalAlignment = xlLeft
            target.VerticalAlignment = xlCenter
            target.WrapText = True
            target.Borders(xlEdgeRight).LineStyle = xlContinuous
        Case 152
            target.Interior.Color = RGB(255,255,255)
            target.Font.Name = "Meiryo UI"
            target.Font.Size = 11
            target.Font.Color = RGB(23,45,74)
            target.HorizontalAlignment = xlLeft
            target.VerticalAlignment = xlCenter
            target.WrapText = True
            target.Borders(xlEdgeLeft).LineStyle = xlContinuous
            target.Borders(xlEdgeBottom).LineStyle = xlContinuous
        Case 153
            target.Interior.Color = RGB(255,255,255)
            target.Font.Name = "Meiryo UI"
            target.Font.Size = 11
            target.Font.Color = RGB(23,45,74)
            target.HorizontalAlignment = xlLeft
            target.VerticalAlignment = xlCenter
            target.WrapText = True
            target.Borders(xlEdgeBottom).LineStyle = xlContinuous
        Case 154
            target.Interior.Color = RGB(255,255,255)
            target.Font.Name = "Meiryo UI"
            target.Font.Size = 11
            target.Font.Color = RGB(23,45,74)
            target.HorizontalAlignment = xlLeft
            target.VerticalAlignment = xlCenter
            target.WrapText = True
            target.Borders(xlEdgeRight).LineStyle = xlContinuous
            target.Borders(xlEdgeBottom).LineStyle = xlContinuous
        Case 155
            target.Interior.Color = RGB(255,255,255)
            target.Font.Name = "Meiryo UI"
            target.Font.Size = 11
            target.Font.Color = RGB(24,48,80)
            target.HorizontalAlignment = xlLeft
            target.VerticalAlignment = xlCenter
            target.WrapText = True
            target.Borders(xlEdgeLeft).LineStyle = xlContinuous
            target.Borders(xlEdgeTop).LineStyle = xlContinuous
            target.Borders(xlEdgeBottom).LineStyle = xlContinuous
        Case 156
            target.Interior.Color = RGB(255,255,255)
            target.Font.Name = "Meiryo UI"
            target.Font.Size = 11
            target.Font.Color = RGB(24,48,80)
            target.HorizontalAlignment = xlLeft
            target.VerticalAlignment = xlCenter
            target.WrapText = True
            target.Borders(xlEdgeTop).LineStyle = xlContinuous
            target.Borders(xlEdgeBottom).LineStyle = xlContinuous
        Case 157
            target.Interior.Color = RGB(255,255,255)
            target.Font.Name = "Meiryo UI"
            target.Font.Size = 11
            target.Font.Color = RGB(24,48,80)
            target.HorizontalAlignment = xlLeft
            target.VerticalAlignment = xlCenter
            target.WrapText = True
            target.Borders(xlEdgeRight).LineStyle = xlContinuous
            target.Borders(xlEdgeTop).LineStyle = xlContinuous
            target.Borders(xlEdgeBottom).LineStyle = xlContinuous
        Case 158
            target.Interior.Color = RGB(255,255,255)
            target.Font.Name = "Meiryo UI"
            target.Font.Size = 11
            target.Font.Color = RGB(24,48,80)
            target.HorizontalAlignment = xlLeft
            target.VerticalAlignment = xlCenter
            target.Borders(xlEdgeLeft).LineStyle = xlContinuous
            target.Borders(xlEdgeRight).LineStyle = xlContinuous
            target.Borders(xlEdgeTop).LineStyle = xlContinuous
            target.Borders(xlEdgeBottom).LineStyle = xlContinuous
        Case 159
            target.Interior.Color = RGB(255,255,255)
            target.Font.Name = "Meiryo UI"
            target.Font.Size = 11
            target.Font.Color = RGB(24,48,80)
            target.HorizontalAlignment = xlLeft
            target.VerticalAlignment = xlCenter
            target.Borders(xlEdgeLeft).LineStyle = xlContinuous
            target.Borders(xlEdgeRight).LineStyle = xlContinuous
            target.Borders(xlEdgeTop).LineStyle = xlContinuous
            target.Borders(xlEdgeBottom).LineStyle = xlContinuous
        Case 160
            target.Interior.Color = RGB(255,255,255)
            target.Font.Name = "Meiryo UI"
            target.Font.Size = 11
            target.Font.Color = RGB(24,48,80)
            target.HorizontalAlignment = xlLeft
            target.VerticalAlignment = xlCenter
            target.WrapText = True
            target.Borders(xlEdgeLeft).LineStyle = xlContinuous
            target.Borders(xlEdgeTop).LineStyle = xlContinuous
            target.Borders(xlEdgeBottom).LineStyle = xlContinuous
        Case 161
            target.Interior.Color = RGB(255,255,255)
            target.Font.Name = "Meiryo UI"
            target.Font.Size = 11
            target.Font.Color = RGB(24,48,80)
            target.HorizontalAlignment = xlLeft
            target.VerticalAlignment = xlCenter
            target.WrapText = True
            target.Borders(xlEdgeRight).LineStyle = xlContinuous
            target.Borders(xlEdgeTop).LineStyle = xlContinuous
            target.Borders(xlEdgeBottom).LineStyle = xlContinuous
        Case 162
            target.Interior.Color = RGB(255,255,255)
            target.Font.Name = "Meiryo UI"
            target.Font.Size = 11
            target.Font.Color = RGB(24,48,80)
            target.HorizontalAlignment = xlLeft
            target.VerticalAlignment = xlCenter
            target.WrapText = True
            target.Borders(xlEdgeLeft).LineStyle = xlContinuous
            target.Borders(xlEdgeTop).LineStyle = xlContinuous
            target.Borders(xlEdgeBottom).LineStyle = xlContinuous
        Case 163
            target.Interior.Color = RGB(255,255,255)
            target.Font.Name = "Meiryo UI"
            target.Font.Size = 11
            target.Font.Color = RGB(24,48,80)
            target.HorizontalAlignment = xlLeft
            target.VerticalAlignment = xlCenter
            target.WrapText = True
            target.Borders(xlEdgeTop).LineStyle = xlContinuous
            target.Borders(xlEdgeBottom).LineStyle = xlContinuous
        Case 164
            target.Interior.Color = RGB(255,255,255)
            target.Font.Name = "Meiryo UI"
            target.Font.Size = 11
            target.Font.Color = RGB(24,48,80)
            target.HorizontalAlignment = xlLeft
            target.VerticalAlignment = xlCenter
            target.WrapText = True
            target.Borders(xlEdgeRight).LineStyle = xlContinuous
            target.Borders(xlEdgeTop).LineStyle = xlContinuous
            target.Borders(xlEdgeBottom).LineStyle = xlContinuous
        Case 165
            target.Interior.Color = RGB(255,255,255)
            target.Font.Name = "Meiryo UI"
            target.Font.Size = 11
            target.Font.Color = RGB(24,48,80)
            target.HorizontalAlignment = xlLeft
            target.VerticalAlignment = xlCenter
            target.Borders(xlEdgeLeft).LineStyle = xlContinuous
            target.Borders(xlEdgeRight).LineStyle = xlContinuous
            target.Borders(xlEdgeTop).LineStyle = xlContinuous
            target.Borders(xlEdgeBottom).LineStyle = xlContinuous
        Case 166
            target.Interior.Color = RGB(255,255,255)
            target.Font.Name = "Meiryo UI"
            target.Font.Size = 11
            target.Font.Color = RGB(24,48,80)
            target.HorizontalAlignment = xlLeft
            target.VerticalAlignment = xlCenter
            target.Borders(xlEdgeLeft).LineStyle = xlContinuous
            target.Borders(xlEdgeRight).LineStyle = xlContinuous
            target.Borders(xlEdgeTop).LineStyle = xlContinuous
            target.Borders(xlEdgeBottom).LineStyle = xlContinuous
        Case 167
            target.Interior.Color = RGB(255,255,255)
            target.Font.Name = "Meiryo UI"
            target.Font.Size = 11
            target.Font.Color = RGB(24,48,80)
            target.HorizontalAlignment = xlLeft
            target.VerticalAlignment = xlCenter
            target.WrapText = True
            target.Borders(xlEdgeLeft).LineStyle = xlContinuous
            target.Borders(xlEdgeTop).LineStyle = xlContinuous
            target.Borders(xlEdgeBottom).LineStyle = xlContinuous
        Case 168
            target.Interior.Color = RGB(255,255,255)
            target.Font.Name = "Meiryo UI"
            target.Font.Size = 11
            target.Font.Color = RGB(24,48,80)
            target.HorizontalAlignment = xlLeft
            target.VerticalAlignment = xlCenter
            target.WrapText = True
            target.Borders(xlEdgeRight).LineStyle = xlContinuous
            target.Borders(xlEdgeTop).LineStyle = xlContinuous
            target.Borders(xlEdgeBottom).LineStyle = xlContinuous
        Case 169
            target.Interior.Color = RGB(242,246,250)
            target.Font.Name = "Meiryo UI"
            target.Font.Size = 9
            target.Font.Color = RGB(23,45,74)
            target.HorizontalAlignment = xlLeft
            target.VerticalAlignment = xlCenter
        Case 170
            target.Interior.Color = RGB(242,246,250)
            target.Font.Name = "Meiryo UI"
            target.Font.Size = 9
            target.Font.Color = RGB(23,45,74)
            target.HorizontalAlignment = xlLeft
            target.VerticalAlignment = xlCenter
        Case 171
            target.Interior.Color = RGB(242,242,242)
            target.Font.Name = "Carlito"
            target.Font.Size = 11
            target.WrapText = True
        Case 172
            target.Interior.Color = RGB(242,242,242)
            target.Font.Name = "Carlito"
            target.Font.Size = 11
            target.WrapText = True
        Case 173
            target.Font.Name = "Meiryo UI"
            target.Font.Size = 9
            target.Font.Color = RGB(100,116,139)
            target.VerticalAlignment = xlCenter
    End Select
End Sub
