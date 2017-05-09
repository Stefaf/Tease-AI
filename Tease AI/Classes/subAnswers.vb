﻿<Serializable>
Public Class subAnswers

	Private checkList As New List(Of String)
	Private answerList As New List(Of String)
	Private ssh As SessionState

	Sub New(session As SessionState)
		ssh = session
		checkList.Add(My.Settings.SubGreeting)
		checkList.Add(My.Settings.SubYes)
		checkList.Add(My.Settings.SubNo)
		checkList.Add(My.Settings.SubSorry)
		checkList.Add("please")
		checkList.Add("thank,thanks")
	End Sub

	Public Function returnWords(s As String) As String
		Select Case s
			Case "hi"
				Return checkList.Item(0)
			Case "yes"
				Return checkList.Item(1)
			Case "no"
				Return checkList.Item(2)
			Case "sorry"
				Return checkList.Item(3)
			Case "thanks"
				Return checkList.Item(5)
			Case "please"
				Return checkList.Item(4)
			Case Else
				Return checkList.Item(0)
		End Select
	End Function

	Public Function returnAll() As List(Of String)
		Return checkList
	End Function

	Public Function isSystemWord(wordList As String) As Boolean
		For i As Integer = 0 To checkList.Count() - 1
			Dim list As String() = ssh.obtainSplitParts(checkList(i), False)
			For n As Integer = 0 To list.Count - 1
				If UCase(wordList).Contains(UCase(list(n))) Then Return True
			Next
		Next
		Return False
	End Function

	Public Sub addToAnswerList(words As String)
		Dim split() = words.Split(",")
		For i As Integer = 0 To split.Count - 1
			answerList.Add(split(i))
		Next
	End Sub

	Public Sub clearAnswers()
		answerList.Clear()
	End Sub

	Public Function triggerWord(chatstring As String) As String
		Dim singleWords() = ssh.obtainSplitParts(chatstring, True)
		For i As Integer = 0 To singleWords.Count - 1
			For n As Integer = 0 To answerList.Count - 1
				If UCase(answerList(n)) = UCase(singleWords(i)) Then Return singleWords(i)
			Next
		Next
		Return ""
	End Function

	Public Function answerNumber() As Integer
		Return answerList.Count
	End Function
End Class
