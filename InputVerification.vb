Module InputVerification
    Public Sub numsOnly(e As KeyPressEventArgs)
        If Asc(e.KeyChar) <> 46 Then
            If Asc(e.KeyChar) <> 8 Then
                If Asc(e.KeyChar) < 48 Or Asc(e.KeyChar) > 57 Then
                    e.Handled = True
                    MsgBox("Numbers Only", MsgBoxStyle.Information, "")
                End If
            End If
        End If
    End Sub
    Public Sub lettersOnly(e As KeyPressEventArgs)
        If AscW(e.KeyChar) > 64 And AscW(e.KeyChar) < 91 Or AscW(e.KeyChar) > 96 And AscW(e.KeyChar) < 123 Or AscW(e.KeyChar) = 8 Then
        Else
            e.KeyChar = Nothing
            MsgBox("Letters Only", MsgBoxStyle.Information, "")
        End If
    End Sub
End Module
