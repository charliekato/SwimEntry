Attribute VB_Name = "gen_entry_form"
'-----
'
'  CAUTION!!! CAUTION!!! CAUTION!!!
'
' Entry sheetのexcelには excelSWMSYSEntrySheet という名前が存在していないと動作しない。
' Entry data のsheetの名前は「申し込み一覧」でないといけない。
'
'


Option Base 1
'
'-----------------------------------------------------
Dim BaseDate As Date
'--------------------------------------------------
'

Dim ShumokuTable(7) As String
Dim DistanceTable(8) As String

Const MAXCOL4HEADER As Integer = 20
Const MAXROW4HEADER As Integer = 1
Const HEADERROW As Integer = 1
Const INDIVENTRYSHEET As String = "個人種目"
Const RELAYENTRYSHEET As String = "リレー種目"
Const SIGNATURE As String = "excelSWMSYSEntrySheet"
Const MAXENTRY As Integer = 2

Const DANTAICODE As String = "25"

Dim ControlWorkbook As Workbook
Dim DataWorkbook As Workbook



Dim ThisFile As String
Dim ListFile As String
Dim ListFilePath As String
'Dim CSVFile As String
Const SWCOMPFile As String = "swcomp.txt"
Const SWENTFile As String = "swent.txt"
Const INFOFile As String = "info.report.txt"
Const SWTEAMFile As String = "swteam.txt"

Dim className(30) As String
Dim TeamNameA(50) As String



Function locate_team_code(teamName As String, init As Boolean)
    Static maxTeamNum As Integer
    Dim id As Integer
    If init Then
        maxTeamNum = 0
        Exit Function
    End If
    For id = 1 To maxTeamNum
        If TeamNameA(id) = teamName Then
            locate_team_code = id
            Exit Function
        End If
    Next id
    TeamNameA(id) = teamName
    maxTeamNum = id
    locate_team_code = id
        
End Function


Sub dump_team()
  Dim num As Integer
  
  For num = 1 To maxTeamNum
    Debug.Print (" " & TeamNameA(num) & " " & num)
  Next
End Sub
Function locate_team_id(teamName As String)
    Dim num As Integer
    For num = 1 To maxTeamNum
        If TeamNameA(num) = teamName Then
            locate_team_id = TeamCodeA(num)
            Exit Function
        End If
    Next
    locate_team_id = 0

End Function

Sub init_style_table()
  ShumokuTable(1) = "自由形"
  ShumokuTable(2) = "背泳ぎ"
  ShumokuTable(3) = "平泳ぎ"
  ShumokuTable(4) = "バタフライ"
  ShumokuTable(5) = "個人メドレー"
  ShumokuTable(6) = "フリーリレー"
  ShumokuTable(7) = "メドレーリレー"
End Sub


Sub init_distance_table()
  DistanceTable(1) = "25m"
  DistanceTable(2) = "50m"
  DistanceTable(3) = "100m"
  DistanceTable(4) = "200m"
  DistanceTable(5) = "300m" ' error
  DistanceTable(6) = "400m"
  DistanceTable(7) = "800m"
  DistanceTable(8) = "1500m"
End Sub


Function locate_style_number(style As String) As Integer
  For locate_style_number = 1 To 7
    If style = ShumokuTable(locate_style_number) Then
      Exit Function
    End If
  Next
End Function
Function locate_distance_code4relay(distanceStr As String) As Integer
    If distanceStr = "100m" Then
        locate_distance_code4relay = 3
    End If
    If distanceStr = "200m" Then
        locate_distance_code4relay = 4
    End If
        If distanceStr = "400m" Then
        locate_distance_code4relay = 5
    End If
        If distanceStr = "800m" Then
        locate_distance_code4relay = 6
    End If
End Function
Function locate_distance_number(distanceStr As String) As Integer
  For locate_distance_number = 1 To 8
    If distanceStr = DistanceTable(locate_distance_number) Then
      Exit Function
    End If
  Next
End Function
    

Function two_byte_style_distance_code(distanceStyle As String, ByRef errorcode As Integer) As String
  Dim mpos As Integer
  Dim styleStr As String
  Dim distanceStr As String
  errorcode = 0
  If distanceStyle = "" Then
    style_distance_code = ""
    Exit Function
  End If
  mpos = InStr(1, distanceStyle, "m")
  If mpos = 0 Then
    style_distance_code = ""
    errorcode = -1
    Exit Function
  End If
  styleStr = Mid(distanceStyle, mpos + 1)
  distanceStr = Left(distanceStyle, mpos)
  two_byte_style_distance_code = CStr(locate_style_number(styleStr)) + _
                                 CStr(locate_distance_number(distanceStr))
End Function

Function style_distance_code_from_str(distanceStyle As String) As String
  Dim mpos As Integer
  Dim styleStr As String
  Dim distanceStr As String
  If distanceStyle = "" Then
    style_distance_code = ""
    Exit Function
  End If
  mpos = InStr(1, distanceStyle, "m")
  If mpos = 0 Then
    style_distance_code = ""
    Exit Function
  End If
  styleStr = Mid(distanceStyle, mpos + 1)
  distanceStr = "00" + Left(distanceStyle, mpos - 1)
  distanceStr = Right(distanceStr, 4)
  style_distance_code_from_str = CStr(locate_style_number(styleStr)) + distanceStr
End Function

Sub styleTest()
  MsgBox (style_distance_code_from_str("50m平泳ぎ"))

End Sub

Function dateFormat(birthday As Date) As String
    Dim ystr As String
    Dim mstr As String
    Dim dstr As String
    ystr = CStr(Year(birthday))
    mstr = Right("0" + CStr(Month(birthday)), 2)
    dstr = Right("0" + CStr(Day(birthday)), 2)
    dateFormat = ystr + mstr + dstr
End Function

Function anotherEntryTimeFormat(entTime As Variant) As String
    Dim tempstr As String
    If entTime = "" Then
        entTime = 999999
    End If
    If entTime > 95999 Then
      tempstr = CStr(entTime)
      anotherEntryTimeFormat = Left(tempstr, 2) + ":" + Mid(tempstr, 3, 2) + "." + Right(tempstr, 2)
    ElseIf entTime > 5999 Then
      tempstr = " " + CStr(entTime)
      anotherEntryTimeFormat = Left(tempstr, 2) + ":" + Mid(tempstr, 3, 2) + "." + Right(tempstr, 2)
    Else
      tempstr = Right("     " + CStr(entTime), 4)
      anotherEntryTimeFormat = "   " + Left(tempstr, 2) + "." + Right(tempstr, 2)
    End If
End Function

Function entryTimeFormat(entTime As Variant) As String
    Dim tempstr As String
    Dim argvartype As Integer
    If entTime = "" Then
        entTime = 999999
    End If
    tempstr = Right("000000" + CStr(entTime), 6)

    entryTimeFormat = Left(tempstr, 4) + "." + Right(tempstr, 2)

End Function



Function column_number(title As String, row As Integer) As Integer
  Dim col As Integer

  For col = 1 To MAXCOL4HEADER

      If StrComp(title, Cells(row, col).Value) = 0 Then
        column_number = col
        Exit Function
      End If

  Next col
  MsgBox (title + " not found... please check the top line of list file")
  column_number = 0
End Function


Function last_row(column As Integer) As Integer
  last_row = Cells(Rows.Count, column).End(xlUp).row
End Function

Function gender_code(genderStr As String) As String
    Dim workStr As String
    workStr = Left(Trim(genderStr), 1)
    If workStr = "男" Then
        gender_code = "1"
    ElseIf workStr = "女" Then
        gender_code = "2"
    Else
        gender_code = "3"
    End If

End Function
Function another_gender_code(genderStr As String) As String
    If genderStr = "男" Then
        another_gender_code = "0"
    Else
        another_gender_code = "5"
    End If
    
    
End Function





Function get_class_number_from_age(myage As Integer) As Integer

    If myage < 30 Then
        get_class_number_from_age = 1
        Exit Function
    End If
    If myage < 40 Then
        get_class_number_from_age = 2
        Exit Function
    End If
    If myage < 50 Then
        get_class_number_from_age = 3
        Exit Function
    End If
    If myage < 60 Then
        get_class_number_from_age = 4
        Exit Function
    End If
    If myage < 70 Then
        get_class_number_from_age = 5
        Exit Function
    End If
    If myage < 80 Then
        get_class_number_from_age = 6
        Exit Function
    End If
    get_class_number_from_age = 7

End Function
Function get_age_from_birthday(birthday As Date) As Integer

    Dim myYear As Integer
    Dim myMonth As Integer
    Dim myDay As Integer
    
    myYear = DateDiff("yyyy", birthday, BaseDate)
    myMonth = Month(birthday)
    myDay = Day(birthday)
    If myMonth < 4 Then
        get_age_from_birthday = myYear
        Exit Function
    End If
    If myMonth = 4 And myDay = 1 Then
        get_age_from_birthday = myYear
        Exit Function
    End If
    get_age_from_birthday = myYear - 1
End Function

Sub age_test(birthday As Date)
    BaseDate = CDate(Range("_baseDate").Value)
    MsgBox (CStr(birthday) + " -> " + CStr(get_age_from_birthday(birthday)))
End Sub
Sub at()
    
    age_test (#4/24/1990#)
    age_test (#9/1/1956#)
End Sub
Function right_padding(orgStr As String, totalLength As Integer) As String

    Dim strLength As Integer
    strLength = LenB(StrConv(orgStr, vbFromUnicode))
    right_padding = orgStr + String(totalLength - strLength, " ")
End Function


Function correct_excel_file() As Boolean
    Dim n As Name
    Dim itemHeader() As Variant
    Dim col As Integer
    itemHeader = Array("No.", "氏名", "ﾌﾘｶﾞﾅ", "所属", "所属カナ", "性別", "生年月日", "学種", "学年", "クラス", "種目1", "エントリータイム1", "種目2")

    correct_excel_file = False
    For col = LBound(itemHeader) To UBound(itemHeader)
    
        If Sheets(INDIVENTRYSHEET).Cells(1, col).Value <> itemHeader(col) Then
            Exit Function
        End If
    Next col
 

    correct_excel_file = True

End Function

Sub read_relay_entry_and_create_swteam(init As Boolean)
    Const startLine As Integer = 2
    Dim lastLine As Integer
    Dim classCode As Integer
    Dim teamNameCol As Integer  'Name of the relay team
    Dim classNameCol As Integer
'    Dim clubNameCol As Integer  'Name of the club to which relay team belongs
    Dim teamNo As String
    Dim teamName As String
    Static idNumber As Integer
    Dim row As Integer
    Dim nameCol As Integer
    Dim genderCol As Integer
    Dim styleCol As Integer
    Dim teamKanaCol As Integer
    Dim style As String
    Dim className As String
    Dim distance As String
    Dim counter As Integer

    Dim timeCol As Integer
    '----- hidden column -------

    Dim outString As String
    Dim shumokuStr As String
    Dim teamKana As String

    Dim teamCode As Integer
    Dim teamID As Integer
    Dim gender As String
    Dim timeStr As String
    Dim clubName As String

    If init Then
        idNumber = 0
        Exit Sub
    End If
    counter = 0

    genderCol = column_number("性別", 1)
    classNameCol = column_number("クラス", 1)
    
    styleCol = column_number("種目", 1)
    timeCol = column_number("エントリータイム", 1)
    teamNameCol = column_number("チーム名", 1)
 '   clubNameCol = column_number("所属", 1)
    teamKanaCol = column_number("ﾌﾘｶﾞﾅ", 1)
    lastLine = last_row(teamNameCol)
    For row = startLine To lastLine

            idNumber = idNumber + 1
            teamName = Cells(row, teamNameCol).Value
  '          clubName = Cells(row, clubNameCol).Value
            teamCode = locate_team_code(teamName, False)
            className = Cells(row, classNameCol).Value
            outString = Right("   " + CStr(idNumber), 4)                                '登録No   1- 4
            outString = outString + right_padding(teamName, 30)                         'チーム名 5-34
            outString = outString + Right("   " & teamCode, 4)                          'teamNo. 35-38
            outString = outString + String(80, " ")                                     '第１泳者-第４泳者(使用せず) 39-118
            outString = outString + DANTAICODE                                          '団体コード 119-120
            outString = outString + "    "                                              '学校コード 121-124  (空白)
            classCode = locate_class_number(className)
            outString = outString + Right("  " & classCode, 2)                          'クラスコード　125-126
            teamKana = Cells(row, teamKanaCol)
            outString = outString + right_padding(teamKana, 15)                         'チーム名カナ　127-141
            style = Cells(row, styleCol).Value
            distance = Left(style, 4)
            style = Right(style, Len(style) - 4)
            outString = outString + _
            right_padding(locate_style_number(style), 16)                               '種目コード    142-157

            outString = outString + right_padding(locate_distance_code4relay(distance), 5)     '距離コード　　158-162
            gender = gender_code(Cells(row, genderCol).Value)                    '
            outString = outString + gender                                              '性別    163
            timeStr = anotherEntryTimeFormat(Cells(row, timeCol).Value)
            
            outString = outString + Right("         " + timeStr, 10)                    'Entry Time   164-173
            outString = outString + String(20, " ")
            counter = counter + 1
            Print #4, outString



    Next row
    Print #3, " リレー " & counter & "種目"
End Sub
Function illegal_class(classNo As Integer, realAge As Integer) As Boolean
    Dim classAge As Integer
    classAge = get_class_number_from_age(realAge)
    If classAge = classNo Then
        illegal_class = False
    Else
        illegal_class = True
    End If

End Function
Function get_gakushu_code(gakushu As String) As Integer
    If gakushu = "小" Then
        get_gakushu_code = 1
        Exit Function
    End If
    If gakushu = "中" Then
        get_gakushu_code = 2
        Exit Function
    End If
    If gakushu = "高" Then
        get_gakushu_code = 3
        Exit Function
    End If
    If gakushu = "大" Then
        get_gakushu_code = 4
        Exit Function
    End If
    If gakushu = "専" Then
        get_gakushu_code = 6
        Exit Function
    End If
    get_gakushu_code = 5
End Function
Sub read_list_file_and_create_swtxt(init As Boolean)
    Const startLine As Integer = 2
    Dim teamNo As String
    Static idNumber As Integer
    Dim row As Integer
    Dim idCol As Integer
    Dim nameCol As Integer
    Dim kanaCol As Integer
    Dim genderCol As Integer
    Dim ageCol As Integer
    Dim birthdayCol As Integer
    Dim teamNameCol As Integer

    Dim teamNoCol As Integer
    Dim gakushuCol As Integer
    Dim gakunenCol As Integer
    Dim styleCol(3) As Integer
    Dim myage As Integer
    Dim className As String ' age class
    Dim myName As String
    Dim timeCol(3) As Integer

    Dim teamCodeCol As Integer

    Dim classCol As Integer
    
    Dim teamNameKanaCol As Integer

    Dim outString As String
    Dim infoString As String
    Dim shumokuStr As String
    Dim shumokuCode As String
    Dim myBirthdayCode As String
    Dim entNo As Integer
    
    Dim teamName As String
    Dim teamKana As String
    Dim gakushu As String
    Dim teamCode As Integer
    Dim teamID As Integer
    Dim gender As String
    Dim genderStr As String

    Dim gakunen As Variant
    Dim gakushuCode As Integer
    
    Dim classNo As Integer
    Dim senshuCount(2) As Integer
    Dim shumokuCount(2) As Integer
    
    Dim errorcode As Integer
    

    If init Then
        idNumber = 0
        Exit Sub
    End If
    senshuCount(1) = 0
    senshuCount(2) = 0
    shumokuCount(1) = 0
    shumokuCount(2) = 0
    idCol = column_number("No.", HEADERROW)
    nameCol = column_number("氏名", HEADERROW)
    kanaCol = column_number("ﾌﾘｶﾞﾅ", HEADERROW)
    genderCol = column_number("性別", HEADERROW)
    ageCol = column_number("クラス", HEADERROW)



    birthdayCol = column_number("生年月日", HEADERROW)
    gakushuCol = column_number("学種", HEADERROW)
    gakunenCol = column_number("学年", HEADERROW)
    styleCol(1) = column_number("種目1", HEADERROW)
    styleCol(2) = column_number("種目2", HEADERROW)
'    styleCol(3) = column_number("種目3", HEADERROW)
    

    timeCol(1) = column_number("エントリータイム1", HEADERROW)
    timeCol(2) = column_number("エントリータイム2", HEADERROW)
'    timeCol(3) = column_number("エントリータイム3", HEADERROW)


    teamNameCol = column_number("所属", HEADERROW)
    teamKanaCol = column_number("所属カナ", HEADERROW)


    For row = startLine To last_row(nameCol)
        myBirthdayCode = Right(dateFormat(Cells(row, birthdayCol).Value), 6)
        idNumber = idNumber + 1
        outString = Right("    " + CStr(idNumber), 5) + "     "                         '登録No  1- 5
        genderStr = Cells(row, genderCol).Value
        gender = gender_code(genderStr)
        outString = outString + gender                                                  '性別    6
        myName = Cells(row, nameCol).Value
        infoString = myName + "  "
        outString = outString + right_padding(myName, 16) + "    "                      '氏名   12-27
        outString = outString + right_padding(Cells(row, kanaCol).Value, 18)            'ﾌﾘｶﾞﾅ  32-49
        outString = outString + String(12, " ")
        teamName = Cells(row, teamNameCol).Value
        teamCode = locate_team_code(teamName, False)

        outString = outString + Right("   " & teamCode, 3)                               'teamNo.62-64
        outString = outString + right_padding(teamName, 16) + "    "                    '所属名1
        teamKana = Cells(row, teamKanaCol)
        outString = outString + right_padding(teamKana, 15) + "     "                   '所属名1　カナ
        outString = outString + String(43, " ")                                         '所属名2 blank
        outString = outString + String(43, " ")                                         '所属名3 blank
        outString = outString + "0"                                                     '使用所属No. 0=大会設定値  191
        
        senshuCount(CInt(gender)) = senshuCount(CInt(gender)) + 1
        If gakushuCol = 0 Then
            gakushuCode = 5
            gakunen = " "
        Else
            gakushuCode = get_gakushu_code(Cells(row, gakushuCol).Value)
            gakunen = Cells(row, gakunenCol).Value
            If gakunen = "" Then
                gakunen = " "
            Else
                gakunen = Cells(row, gakunenCol).Value
            End If
        End If
        outString = outString & gakushuCode                                                '学校code 192 5=一般

        outString = outString & gakunen                                                 '学年　(一般はブランク) 193
        outString = outString & myBirthdayCode                                          '生年月日  194-199
        myage = get_age_from_birthday(CDate(Cells(row, birthdayCol).Value))
        outString = outString & Right(" " + CStr(myage), 2)                             '年齢　200-201

        outString = outString & "    "                                                  '予備 202-205
        outString = outString & DANTAICODE                                              '団体コード=25  206-207
        outString = outString + DANTAICODE & "001" & myBirthdayCode                  '旧日水連コード 25001560901
'        outString = outString & "           "
        outString = outString & another_gender_code(genderStr)
'        outString = outString & "     "
        outString = outString & String(37, " ")
        Print #1, outString

        className = Cells(row, ageCol).Value
        classNo = locate_class_number(className)
 '       If illegal_class(classNo, myage) Then
 '           Print #3, "(W) 基準日年齢とクラスが合いません。" & myName & "　基準日年齢 : " & myage & "  クラス : " & classname
 '       End If
        For entNo = 1 To MAXENTRY
            shumokuStr = Cells(row, styleCol(entNo)).Value
            shumokuCode = two_byte_style_distance_code(shumokuStr, errorcode)
            If shumokuCode <> "" Then
                infoString = infoString + shumokuStr + "  "
                outString = Right("    " + CStr(idNumber), 5) + "     "
                outString = outString & Right(" " + CStr(entNo), 2)
                outString = outString & gender
                outString = outString & shumokuCode
                outString = outString & Right("  " & classNo, 2)
                outString = outString & anotherEntryTimeFormat(Cells(row, timeCol(entNo)).Value) + " "
                outString = outString & "      "
'                Debug.Print (" swent length is " + CStr(LenB(StrConv(outString, vbFromUnicode))))
                Print #2, outString
                shumokuCount(gender) = shumokuCount(gender) + 1
            ElseIf errorcode < 0 Then
                Print #3, ""
                Print #3, "(Error!) 不正な種目です　" & shumokuStr & "  プルダウンメニューから選んでいないと思われます。"
            End If
        Next entNo
        Print #3, infoString

    Next row
    Print #3, "*** summary ***"
    Print #3, "女子 : " & senshuCount(2) & "名  " & shumokuCount(2) & "種目"
    Print #3, "男子 : " & senshuCount(1) & "名  " & shumokuCount(1) & "種目"

End Sub
'obsolete
Sub dcteat()
    debugclass ("18～24歳")
End Sub
Function locate_class_number(cname As String) As Integer
    Dim i As Integer
    Dim rng As Range
    locate_class_number = Application.WorksheetFunction.VLookup(cname, Range("classTable"), 2, False)
  
End Function
Sub debugclass(className As String)
    MsgBox (className & " " & locate_class_number(className))
End Sub

Sub create_go()
    Dim buf As String
    Dim inputFilePath As String
    Dim outputFilePath As String
    Call init
    
    Call read_list_file_and_create_swtxt(True)
    Call read_relay_entry_and_create_swteam(True)
    Call locate_team_code("", True)
    inputFilePath = ""

    inputFilePath = get_folder("申し込みのエクセルファイル(入力ファイル)があるフォルダを選んでください", "C:\")
    'inputFilePath = get_folder2()
    If inputFilePath = "" Then Exit Sub
    
    outputFilePath = get_folder("競泳リザルトシステムに渡すファイル(出力ファイル)を入れるフォルダを選んでください", "C:\")
    If outputFilePath = "" Then Exit Sub
    Open outputFilePath + "\" + SWCOMPFile For Output As #1
    Open outputFilePath + "\" + SWENTFile For Output As #2
    Open outputFilePath + "\" + INFOFile For Output As #3
    Open outputFilePath + "\" + SWTEAMFile For Output As #4
    buf = Dir(inputFilePath + "\*.xls?")

    Do While Len(buf) > 0
        Call openListFile(inputFilePath, buf)
        If correct_excel_file() Then
            Print #3, "Processing " & buf & "..."
            Sheets(INDIVENTRYSHEET).Select
            If data_found() Then
                Call read_list_file_and_create_swtxt(False)
            Else
                Print #3, " Data not found."
            End If
            Sheets(RELAYENTRYSHEET).Select
            Call read_relay_entry_and_create_swteam(False)
            Print #3, ""
            Print #3, ""
        Else
            Print #3, "" & buf & " is not the correct entry sheet. Skipping it..."
        End If

        Call closeListFile(buf)
        buf = Dir()
    Loop
    Close #1
    Close #2
    Close #3
    Close #4
'    ControlWorkbook.Sheets("Sheet1").Select
    MsgBox ("完了しました。")
End Sub
Sub write_info(path As String, filename As String)
    Dim gender, j As Integer
    Open path + "\" + filename For Output As #1
    
    For j = 1 To 16
        Print #1, "**** " & Shozoku_array(j) & "****"
        Print #1, "個人種目参加者"
        Print #1, "    男子 : " & Right("  " & Senshu_count(1, j), 3)
        Print #1, "    女子 : " & Right("  " & Senshu_count(2, j), 3)
    Next j
    Close #1
End Sub
Function data_found()
    If Cells(3, 2).Value = "" Then
        data_found = False
    Else
        data_found = True
    End If
End Function
Sub init()
'    BaseDate = CDate(Range("_baseDate").Value)

    Call init_style_table
    Call init_distance_table
    maxTeamNum = 0


End Sub

Sub ctest()
    Windows("スポーツマスターズ参加申込書.xlsx").Activate
    If correct_excel_file() Then
        MsgBox ("OK")
    End If
End Sub
Sub openListFile(pathname As String, filename As String)

    ThisFile = ActiveWorkbook.Name
    Workbooks.Open filename:=pathname + "\" + filename
    Windows(filename).Activate

End Sub



Sub closeListFile(filename)

  Windows(filename).Activate
  ActiveWorkbook.Close savechanges:=False
End Sub
Function get_folder(message As String, initFolder As String) As String
    get_folder = ""
    Dim Shell, myPath
    Set Shell = CreateObject("Shell.Application")
    Set myPath = Shell.BrowseForFolder(&O0, message, &H1 + &H10, initFolder)
    If Not myPath Is Nothing Then get_folder = myPath.Items.Item.path
    Set Shell = Nothing
    Set myPath = Nothing
End Function

Function get_folder2() As String

    With Application.FileDialog(msoFileDialogFolderPicker)
        If .Show = True Then
            get_folder2 = .SelectedItems(1)
        End If
    End With
End Function
