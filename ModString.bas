Attribute VB_Name = "ModString"
Option Explicit

Public Function GetString(strdata As String, sFind As String, nData() As String, Optional NoString As Boolean, Optional GetDataStringCount As Integer, Optional OnsFind As String) As String
    Dim SInStr As Long, tmpSInStr As Long, nSInStr As Long, offnSInStr As Boolean
    Dim MnData As Integer, dsss As Integer, offLoad As Boolean
    Dim tmpGetString As String
    
    ReDim nData(MnData)
    Do
        If offLoad = False Then tmpSInStr = SInStr
        If offnSInStr = False Then
            nSInStr = InStr(nSInStr + 1, strdata, "'")
            If nSInStr <> 0 Then
                dsss = 1 + -dsss
                offnSInStr = True
                If dsss = 0 Then
                    offLoad = False
                    SInStr = nSInStr - 0
                    dsss = 0
                End If
            Else
                offnSInStr = True
            End If
        End If
        
        If offLoad = False Then
            SInStr = InStr(SInStr + 1, strdata, sFind)
        End If
        
        If SInStr > nSInStr And nSInStr <> 0 Then
            offnSInStr = False
            offLoad = True
        Else
            offLoad = False
        End If
        
        If offLoad = False Then
            If SInStr <> 0 Then
                ReDim Preserve nData(MnData)
                nData(MnData) = Right(Left(strdata, SInStr - 1), Abs(SInStr - tmpSInStr) - 1)
                If GetDataStringCount = -1 Then
                    GetString = tmpGetString
                    tmpGetString = tmpGetString & nData(MnData) & OnsFind
                Else
                    If MnData < GetDataStringCount Then GetString = GetString & nData(MnData) & OnsFind
                End If
                MnData = MnData + 1
            Else
                Exit Do
            End If
        End If
    Loop
End Function

