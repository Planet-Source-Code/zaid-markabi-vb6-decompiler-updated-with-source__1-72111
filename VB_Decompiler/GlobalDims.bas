Attribute VB_Name = "GlobalDims"
' VB Decompiler Lite
' Programmed By [ Zaid Markabi ]
' ___________________________________________________________________________________________________
'|                                                                                                   |\_______________________
'|  ###############        ###         #####   ######                ######    #####                 |                        |\0 1 1 1 0 0 1 1 0 0 0 1 0 0 1 0 0 1 0 0 1 1 1 1 0 0 1 1 0 0 0 1 0 0 1 0 0 1 0 0 1
'| ##############         #####         ###     ##   ##               ######  #####                  |      Zaid Markabi      |=\ 1 0 0 1 0 0 0 0 0 1 1 0 1 0 0 0 1 1 1 0 1 0 0 1 0 0 0 0 0 1 1 0 1 0 0 0 1 1 1 0
'|         ####          ### ###        ###     ##    ##              ##  ## ##  ##                  |                        |==\0 0 1 1 1 0 1 0 0 1 0 0 1 1 0 0 1 0 1 1 0 0 1 1 1 0 1 0 0 1 0 0 1 1 0 0 1 0 1 1
'|       ###            ###   ###       ###     ##     ##    #####    ##   ###   ##                  | zaidmarkabi@yahoo.com  |===\ 1 __________________________________  0 1 0 0 0 1 1 1 0 1 0 0 1 0 0 1 0 0 0 1
'|     ###             ###########      ###     ##     ##   ####      ##    #    ##                  |                        |====|>| Development For Our Digital Life | 1 1 0 0 1 1 1 0 1 0 0 1 0 0 0 1 1 0 1 0
'|   ###              #############     ###     ##    ##              ##         ##      A R K A B I | VisualBasic Programmer |===/ 1|__________________________________| 0 1 1 0 1 0 0 0 1 1 1 0 1 0 1 1 0 1 0 0
'| ##############    ###         ###    ###     ##   ##               ##         ##     ############ |                        |==/0 0 1 1 1 0 1 0 0 1 0 0 1 1 0 0 1 0 1 1 0 0 1 1 1 0 1 0 0 1 0 0 1 1 0 0 1 0 1 1
'| ###############   ###         ###   #####   ######                ####       ####   ### 2009 ###  |Syria ( Arabic )-Tartous|=/ 1 0 0 1 0 0 0 0 0 1 1 0 1 0 0 0 1 1 1 0 1 0 0 1 0 0 0 0 0 1 1 0 1 0 0 0 1 1 1 0
'|                                                                                    ############   | _______________________|/0 1 1 1 0 0 1 1 0 0 0 1 0 0 1 0 0 1 0 0 1 1 1 1 0 0 1 1 0 0 0 1 0 0 1 0 0 1 0 0 1
'|___________________________________________________________________________________________________|/
'
' Em@il    : zaidmarkabi@yahoo.com
' Web site : www.YazanMarkabi.Webs.com
'            I hope to hear from you ,

Global FileName As String
Global FilePath As String

Global CHAR As Byte
Global CHARn() As Byte
Global FilePos As Long
Global LastSentenceText As String
Global LastSavedSentence As String
Global MaxSentenceLong As Integer

Function LoadImage(FileName) As Boolean
On Error GoTo Err
LoadImage = False
frmExtractingData.LastImage.Picture = LoadPicture(FileName)
LoadImage = True
Exit Function
Err:
End Function


Function IfItIsLettel(StrText As String) As Boolean
If InStr(1, "AaBbCcDdEeFfGgHhIiJjKkLlMmNnOoPpQqRrSsTtUuVvWwXxYyZz :-_0123456789", StrText) > 0 Then
IfItIsLettel = True
Else
IfItIsLettel = False
End If
End Function


Sub ExtractNext()
Dim I As Integer

If EOF(1) = False Then

Dim FileHeader As String
Dim FileType As String
Dim HeaderLong As Long

' Get the next header ( 3 charts is enough )
For I = 0 To 2
Get #1, FilePos + I, CHAR
FileHeader = FileHeader + Chr(CHAR)
Next
Get #1, FilePos, CHAR

If frmExtractingData.ChkImages.Value = 1 Then
' Define types of images ( Headers )
If Left(UCase(FileHeader), 2) = "BM" Then
FileType = "Bmp"
HeaderLong = 0
End If
If UCase(FileHeader) = "JFI" Then
FileType = "Jpg"
HeaderLong = 6
FilePos = FilePos - 6
End If
' Gif and Png in FULL version

'  Extracting Images ( Jpeg , Bmp , Gif , Png )
If FileType = "Bmp" Or FileType = "Jpg" Or FileType = "Gif" Or FileType = "Png" Then
ReDim CHARn(LOF(1) - FilePos)
Get #1, FilePos, CHARn
Open App.Path + "\Temp\" + Mid(FileName, Len(FilePath) + 1, Len(FileName) - Len(FilePath) - 4) + Format(FilePos, "000000000000000") + "." + FileType For Binary As #2
Put #2, 1, CHARn
Close #2
If LoadImage(App.Path + "\Temp\" + Mid(FileName, Len(FilePath) + 1, Len(FileName) - Len(FilePath) - 4) + Format(FilePos, "000000000000000") + "." + FileType) = True Then
If FileType = "Bmp" Then
SavePicture frmExtractingData.LastImage.Picture, App.Path + "\Temp\" + Mid(FileName, Len(FilePath) + 1, Len(FileName) - Len(FilePath) - 4) + Format(FilePos, "000000000000000") + "." + FileType
End If
Open App.Path + "\Temp\" + Mid(FileName, Len(FilePath) + 1, Len(FileName) - Len(FilePath) - 4) + Format(FilePos, "000000000000000") + "." + FileType For Binary As #3
HeaderLong = LOF(3)
Close #3
GoTo ImgEx
Else
Kill App.Path + "\Temp\" + Mid(FileName, Len(FilePath) + 1, Len(FileName) - Len(FilePath) - 4) + Format(FilePos, "000000000000000") + "." + FileType
End If
End If
End If

If frmExtractingData.ChkSentences.Value = 1 Then
' Extract Sentences
If IfItIsLettel(Chr(CHAR)) = True Then
LastSentenceText = LastSentenceText + Chr(CHAR)
Else
If Len(LastSentenceText) > MaxSentenceLong Then
Open App.Path + "\Temp\" + Mid(FileName, Len(FilePath) + 1, Len(FileName) - Len(FilePath) - 4) + Format(FilePos - Len(LastSentenceText), "000000000000000") + ".Txt" For Binary As #8
Put #8, 1, LastSentenceText
Close #8
LastSavedSentence = LastSentenceText
LastSentenceText = ""
End If
End If
End If

ImgEx:
FilePos = FilePos + 1 + HeaderLong

Else
frmExtractingData.TimerForExtract.Enabled = False
End If
End Sub

Sub RefreshLoadingBar()
On Error GoTo Done
frmExtractingData.LoadingExtDone.Width = (frmExtractingData.LoadingExtAll.ScaleWidth * FilePos) / LOF(1)
If EOF(1) = True Then
MsgBox vbCrLf + vbCrLf + Space(20) + "Done !" + Space(30) + vbCrLf + vbCrLf, vbInformation, "Finished !"
Load frmEditor
frmExtractingData.RefList.Enabled = False
frmExtractingData.TimerForExtract.Enabled = False
frmExtractingData.Hide
Close #1
End If
frmExtractingData.Label1(3).Caption = "Extracted Sentences " + Format(frmExtractingData.ExtractedFilesSentences.ListCount)
frmExtractingData.Label1(0).Caption = "Extracted Images " + Format(frmExtractingData.ExtractedFiles.ListCount)
Exit Sub
Done:
frmExtractingData.LoadingExtDone.Width = frmExtractingData.LoadingExtAll.ScaleWidth
End Sub

Sub RefreshLoadingBarB(BarPro As Integer)
On Error GoTo Done
frmEditor.LoadingExtDone.Width = (frmEditor.LoadingExtAll.ScaleWidth * BarPro) / frmEditor.ExtractedFilesSentences.ListCount
Exit Sub
Done:
frmEditor.LoadingExtDone.Width = frmEditor.LoadingExtAll.ScaleWidth
End Sub
