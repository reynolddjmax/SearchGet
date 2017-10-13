

''' <summary>
''' '''''''''''''''''''''''''''''''
''' </summary>
Public Class ID

    '身份证检查
    Shared Function IDCK(ByVal IdCard As String) As String '身份证号检查

        IDDate = "1000-1-1"
        SexCode = 3


        Dim sDate As String
        sDate = ""

        '长度+数字格式检查
        If IdCard = "" Then
            IDCK = "身份证为空白"
            Exit Function
        ElseIf IdCard Like "[1-9]#################" = False And IdCard Like "[1-9]################X" = False And IdCard Like "[1-9]################x" = False Then 'And IdCard Like "[1-9]##############" = False 
            IDCK = "身份证长度格式错误"
            Exit Function
        End If


        '出生日期
        If Len(IdCard) = 15 Then
            sDate = Mid(IdCard, 7, 2) & "-" & Mid(IdCard, 9, 2) & "-" & Mid(IdCard, 11, 2)
            SexCode = Strings.Right(IdCard, 1) Mod 2
        ElseIf Len(IdCard) = 18 Then
            sDate = Mid(IdCard, 7, 4) & "-" & Mid(IdCard, 11, 2) & "-" & Mid(IdCard, 13, 2)
            SexCode = Mid(IdCard, 17, 1) Mod 2
            '检查检验码
            Dim CheckCodeCacu As String
            CheckCodeCacu = GetCheckCode(CStr(IdCard))   '计算检验码
            If CheckCodeCacu <> LCase(Strings.Right(IdCard, 1)) Then
                IDCK = "身份证错误：检验码检查未通过"
                Exit Function
            End If
        Else
            MsgBox("Error 0001")
        End If

        If IsDate(sDate) = False Then
            IDCK = "身份证日期格式错误"
            Exit Function
        End If


        IDDate = sDate

        '寿命检查
        Dim life As Integer
        life = DateDiff(DateInterval.Month, IDDate, Now)
        If life > 1560 Then
            IDCK = "身份证错误：年龄过大"
            Exit Function
        End If

        If life < 2 Then
            IDCK = "身份证错误：年龄过小"
            Exit Function
        End If

        IDCK = ""

    End Function


    Shared Function GetCheckCode(ByVal IdCard As String) As String '计算检校码
        Dim i As Integer
        Dim ll_code(17) As Long
        Dim ll_sum, li_number As Long
        Dim Is_checkcode As String
        For i = 1 To 17
            '得到加权因子值  wi=2 ^(i-1) mod 11 [i 18 -2 ]
            ll_code(i) = (2 ^ (18 - i)) Mod 11
            li_number = Val(Mid(IdCard, i, 1))
            ll_sum = ll_sum + ll_code(i) * li_number

        Next i

        Is_checkcode = Trim(Str((ll_sum Mod 11)))

        Select Case Is_checkcode
            Case "2"
                GetCheckCode = "X"
            Case "0", "1"
                GetCheckCode = CStr(0 ^ Int(Is_checkcode))
            Case Else
                GetCheckCode = CStr(12 - Val(Trim(Is_checkcode)))
        End Select
    End Function

End Class
