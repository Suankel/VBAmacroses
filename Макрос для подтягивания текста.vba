Sub lift_up_lines()

Dim MyCell As Range
    Dim stringRanges As String
      stringRanges = "C79:C98,G79:G98,K79:K98,O79:O98,S79:S98,C111:C130,G111:G130,K111:K130,O111:O130,S111:S130,C137:C156,G137:G156,K137:K156,O137:O156,S137:S156," _
    & "C168:C187,G168:G187,K168:K187,O168:O187,S168:S187,C193:C212,G193:G212,K193:K212,O193:O212,S193:S212,C214:C233,G214:G233,K214:K233,O214:O233,S214:S233," _
    & "C234:C253,G234:G253,K234:K253,O234:O253,S234:S253,C254:C273,G254:G273,K254:K273,O254:O273,S254:S273,C274:C293,G274:G293,K274:K293,O274:O293,S274:S293," _
    & "C294:C313,G294:G313,K294:K313,O294:O313,S294:S313,C314:C333,G314:G333,K314:K333,O314:O333,S314:S333,C334:C353,G334:G353,K334:K353,O334:O353,S334:S353," _
    & "C354:C373,G354:G373,K354:K373,O354:O373,S354:S373,C375:C394,G375:G394,K375:K394,O375:O394,S375:S394,C396:C415,G396:G415,K396:K415,O396:O415,S396:S415," _
    & "C416:C435,G416:G435,K416:K435,O416:O435,S416:S435,C436:C455,G436:G455,K436:K455,O436:O455,S436:S455,C456:C475,G456:G475,K456:K475,O456:O475,S456:S475," _
    & "C476:C495,G476:G495,K476:K495,O476:O495,S476:S495,C496:C515,G496:G515,K496:K515,O496:O515,S496:S515,C516:C535,G516:G535,K516:K535,O516:O535,S516:S535," _
    & "C536:C555,G536:G555,K536:K555,O536:O555,S536:S555,C566:C585,G566:G585,K566:K585,O566:O585,S566:S585,C596:C615,G596:G615,K596:K615,O596:O615,S596:S615," _
    & "C621:C640,G621:G640,K621:K640,O621:O640,S621:S640,C647:C666,G647:G666,K647:K666,O647:O666,S647:S666,C668:C687,G668:G687,K668:K687,O668:O687,S668:S687," _
    & "C688:C707,G688:G707,K688:K707,O688:O707,S688:S707,C716:C735,G716:G735,K716:K735,O716:O735,S716:S735,C745:C764,G745:G764,K745:K764,O745:O764,S745:S764," _
    & "C773:C792,G773:G792,K773:K792,O773:O792,S773:S792,C801:C820,G801:G820,K801:K820,O801:O820,S801:S820,C829:C848,G829:G848,K829:K848,O829:O848,S829:S848," _
    & "C857:C876,G857:G876,K857:K876,O857:O876,S857:S876,C885:C904,G885:G904,K885:K904,O885:O904,S885:S904,C913:C932,G913:G932,K913:K932,O913:O932,S913:S932," _
    & "C941:C960,G941:G960,K941:K960,O941:O960,S941:S960,C969:C988,G969:G988,K969:K988,O969:O988,S969:S988,C997:C1016,G997:G1016,K997:K1016,O997:O1016,S997:S1016," _
    & "C1023:C1042,G1023:G1042,K1023:K1042,O1023:O1042,S1023:S1042,C1049:C1068,G1049:G1068,K1049:K1068,O1049:O1068,S1049:S1068,C1075:C1094,G1075:G1094,K1075:K1094,O1075:O1094,S1075:S1094," _
    & "C1096:C1115,G1096:G1115,K1096:K1115,O1096:O1115,S1096:S1115,C1117:C1136,G1117:G1136,K1117:K1136,O1117:O1136,S1117:S1136,C1138:C1157,G1138:G1157,K1138:K1157,O1138:O1157,S1138:S1157"
 ' you can choose another range
    
     Dim arrRanges() As String
    Dim stringCommon As String
    arrRanges = Split(stringRanges, ",")
    Dim stringForLoop As Variant
    Dim arr(0 To 19) As String
    Dim counter As Integer
    Dim countNoZeroCells As Integer
    Dim myReg2 As New RegExp
    Dim myReg3 As New RegExp
    Dim myStr As String
    
    Set objRegExp = CreateObject("VBScript.RegExp")
    objRegExp.Global = True
    objRegExp.IgnoreCase = True
    objRegExp.MultiLine = True
    objRegExp.Pattern = "()" 'you can write phrases separated by the "|" that you need to convert to uppercase
    
    myReg2.Pattern = "\.\s*$"
    myReg3.Pattern = "^\,\s\,\s\,\s\n\,\s\,\s\,\s$"
    
    countNoZeroCells = 0
    
    For Each stringForLoop In arrRanges
        Set MyRange = Range(stringForLoop)
        counter = 0
        For Each MyCell In MyRange
                myStr = MyCell.Value
                If myReg2.Test(myStr) Then
                    MsgBox (myStr)
                    MyCell.Value = myReg2.Replace(myStr, "")
                End If
                If myReg3.Test(myStr) Then
                    MyCell.Value = myReg3.Replace(myStr, "")
                End If
                
                Set objMatches = objRegExp.Execute(myStr)
                If objMatches.Count <> 0 Then
                    For Each ret In objMatches
                        MyCell.Value = Replace(myStr, ret, UCase(ret), 1, 1, vbTextCompare)
   
                    Next ret
                End If
                       
            If MyCell.Value <> 0 Then
                arr(counter) = MyCell.Value
                counter = counter + 1
                countNoZeroCells = counter
            Else
                MyCell.Value = ""
            End If
        Next MyCell
        
        counter = 0
        For Each MyCell In MyRange
            If countNoZeroCells = 0 Then
                MyRange.Cells(1, 1) = "-"
            Else
                MyCell.Value = Replace(arr(counter), Left(arr(counter), 1), UCase(Left(arr(counter), 1)), 1, 1, 1)
                counter = counter + 1
                If (counter = 19) Then
                    counter = 0
                    Erase arr
                End If
            End If
        Next MyCell
        countNoZeroCells = 0
    Next stringForLoop
MsgBox ("All lines are lifted up")

End Sub

