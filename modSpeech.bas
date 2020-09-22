Attribute VB_Name = "modSpeech"
'*************************************************************************
'
' modSpeech for numeric and sign (#%$.) speech
'
' Written by:  Blake B. Pell
'              blakepell@hotmail.com, bpell@indiana.edu
'
' This code is free for personal/freeware use, any commerical use needs
' the consent of the author (Blake Pell).
'
' This module contains the following Subs and/or Functions:
'
' Sub PlayWave
' Sub NumericParseTextForSound(myText as String)
' Sub PlayWaveOver()
' Function Initialize_Audio() as Boolean
'
'*************************************************************************


Public Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long

Public Sub PlayWave(filename As String)
' PlayWave WILL wait for the first wave to end
x = sndPlaySound(filename, 2)
End Sub

Public Sub PlayWaveOver(filename As String)
' PlayWaveOver won't wait for the first wave to end
x = sndPlaySound(filename, 1)
End Sub

Public Sub NumericParseTextForSound(myText As String)

DoEvents

' Find case, will use function over and over until string is parsed

If Len(myText) = 1 Then


        If Mid$(myText, 1, 1) = "1" Then
            PlayWave ("one.wav")
        ElseIf Mid$(myText, 1, 1) = "2" Then
            PlayWave ("two.wav")
        ElseIf Mid$(myText, 1, 1) = "3" Then
            PlayWave ("three.wav")
        ElseIf Mid$(myText, 1, 1) = "4" Then
            PlayWave ("four.wav")
        ElseIf Mid$(myText, 1, 1) = "5" Then
            PlayWave ("five.wav")
        ElseIf Mid$(myText, 1, 1) = "6" Then
            PlayWave ("six.wav")
        ElseIf Mid$(myText, 1, 1) = "7" Then
            PlayWave ("seven.wav")
        ElseIf Mid$(myText, 1, 1) = "8" Then
            PlayWave ("eight.wav")
        ElseIf Mid$(myText, 1, 1) = "9" Then
            PlayWave ("nine.wav")
        ElseIf Mid$(myText, 1, 1) = "0" Then
            ' PlayWave ("zero.wav")
        End If

    
ElseIf Len(myText) = 2 Then
    
        If Mid$(myText, 1, 1) = "1" Then
            ' 10 - 12
            If Mid$(myText, 2, 1) = "1" Then
                PlayWave ("eleven.wav")
            ElseIf Mid$(myText, 2, 1) = "2" Then
                PlayWave ("twelve.wav")
            ElseIf Mid$(myText, 2, 1) = "3" Then
                PlayWave ("thirteen.wav")
            ElseIf Mid$(myText, 2, 1) = "4" Then
                PlayWave ("fourteen.wav")
            ElseIf Mid$(myText, 2, 1) = "5" Then
                PlayWave ("fifteen.wav")
            ElseIf Mid$(myText, 2, 1) = "6" Then
                PlayWave ("sixteen.wav")
            ElseIf Mid$(myText, 2, 1) = "7" Then
                PlayWave ("seventeen.wav")
            ElseIf Mid$(myText, 2, 1) = "8" Then
                PlayWave ("eightteen.wav")
            ElseIf Mid$(myText, 2, 1) = "9" Then
                PlayWave ("nineteen.wav")
            ElseIf Mid$(myText, 2, 1) = "0" Then
                PlayWave ("ten.wav")
            End If
        ElseIf Mid$(myText, 1, 1) = "2" Then
            PlayWave ("twenty.wav")
            NumericParseTextForSound (Mid$(myText, 2, 1))
        ElseIf Mid$(myText, 1, 1) = "3" Then
            PlayWave ("thirty.wav")
            NumericParseTextForSound (Mid$(myText, 2, 1))
        ElseIf Mid$(myText, 1, 1) = "4" Then
            PlayWave ("fourty.wav")
            NumericParseTextForSound (Mid$(myText, 2, 1))
        ElseIf Mid$(myText, 1, 1) = "5" Then
            PlayWave ("fifty.wav")
            NumericParseTextForSound (Mid$(myText, 2, 1))
        ElseIf Mid$(myText, 1, 1) = "6" Then
            PlayWave ("sixty.wav")
            NumericParseTextForSound (Mid$(myText, 2, 1))
        ElseIf Mid$(myText, 1, 1) = "7" Then
            PlayWave ("seventy.wav")
            NumericParseTextForSound (Mid$(myText, 2, 1))
        ElseIf Mid$(myText, 1, 1) = "8" Then
            PlayWave ("eighty.wav")
            NumericParseTextForSound (Mid$(myText, 2, 1))
        ElseIf Mid$(myText, 1, 1) = "9" Then
            PlayWave ("ninety.wav")
            NumericParseTextForSound (Mid$(myText, 2, 1))
        ElseIf Mid$(myText, 1, 1) = "0" Then
            NumericParseTextForSound (Mid$(myText, 2, 1))
        End If
    
ElseIf Len(myText) = 3 Then
    ' hundreds
    ' 000's
    If Mid$(myText, 1, 1) <> "0" Then
        NumericParseTextForSound (Mid$(myText, 1, 1))
        PlayWave ("hundred.wav")
    End If
    
    NumericParseTextForSound (Mid$(myText, 2, 2))
    
ElseIf Len(myText) = 4 Then
    ' thousands
    NumericParseTextForSound (Mid$(myText, 1, 1))
    PlayWave ("thousand.wav")
    NumericParseTextForSound (Mid$(myText, 2, 4))
ElseIf Len(myText) = 5 Then
    ' ten thousands
    NumericParseTextForSound (Mid$(myText, 1, 2))
    PlayWave ("thousand.wav")
    NumericParseTextForSound (Mid$(myText, 3, 5))
ElseIf Len(myText) = 6 Then
    ' hundred thousands
    NumericParseTextForSound (Mid$(myText, 1, 3))
    PlayWave ("thousand.wav")
    NumericParseTextForSound (Mid$(myText, 4, 3))
ElseIf Len(myText) = 7 Then
    ' millions
    NumericParseTextForSound (Mid$(myText, 1, 1))
    PlayWave ("million.wav")
    NumericParseTextForSound (Mid$(myText, 2, 7))
ElseIf Len(myText) = 8 Then
    ' 10 millions
    NumericParseTextForSound (Mid$(myText, 1, 2))
    PlayWave ("million.wav")
    NumericParseTextForSound (Mid$(myText, 3, 8))
ElseIf Len(myText) = 9 Then
    ' hundred millions
    NumericParseTextForSound (Mid$(myText, 1, 3))
    PlayWave ("million.wav")
    NumericParseTextForSound (Mid$(myText, 4, 9))
ElseIf Len(myText) = 10 Then
    ' billions
    NumericParseTextForSound (Mid$(myText, 1, 1))
    PlayWave ("billion.wav")
    NumericParseTextForSound (Mid$(myText, 2, 10))
ElseIf Len(myText) = 11 Then
    ' 10 billions
    NumericParseTextForSound (Mid$(myText, 1, 2))
    PlayWave ("billion.wav")
    NumericParseTextForSound (Mid$(myText, 3, 11))
ElseIf Len(myText) = 12 Then
    ' 100 billions
    NumericParseTextForSound (Mid$(myText, 1, 3))
    PlayWave ("billion.wav")
    NumericParseTextForSound (Mid$(myText, 4, 12))
ElseIf Len(myText) = 13 Then
    ' trillions
    NumericParseTextForSound (Mid$(myText, 1, 1))
    PlayWave ("trillion.wav")
    NumericParseTextForSound (Mid$(myText, 2, 13))
ElseIf Len(myText) = 14 Then
    ' 10 trillions
    NumericParseTextForSound (Mid$(myText, 1, 2))
    PlayWave ("trillion.wav")
    NumericParseTextForSound (Mid$(myText, 3, 14))
ElseIf Len(myText) = 15 Then
    ' 100 trillions
    NumericParseTextForSound (Mid$(myText, 1, 3))
    PlayWave ("trillion.wav")
    NumericParseTextForSound (Mid$(myText, 4, 15))
End If


End Sub

Public Sub SymbolParseTextForSound(myText As String)
    
    If Len(myText) <> 1 Then Exit Sub
    
    If myText = "$" Then
        PlayWave ("dollars.wav")
    ElseIf myText = "." Then
        PlayWave ("dot.wav")
    ElseIf myText = "%" Then
        PlayWave ("percent.wav")
    ElseIf myText = "#" Then
        PlayWave ("number.wav")
    ElseIf myText = "&" Then
        PlayWave ("cents.wav")
    End If
    
End Sub

