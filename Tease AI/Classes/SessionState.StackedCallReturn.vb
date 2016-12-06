Partial Class SessionState
	'test lineklkl
	Friend Class StackedCallReturn
		Dim numReturns As Integer = -1
		Dim ReturnStrokeTauntVal As List(Of Integer) = New List(Of Integer)
		Dim ReturnFileText As List(Of String) = New List(Of String)
		Dim ReturnState As List(Of String) = New List(Of String)

		Public Sub AddFile(ByVal _Session As SessionState)
			ReturnFileText.Add(_Session.FileText)
			ReturnStrokeTauntVal.Add(_Session.StrokeTauntVal)
			ReturnState.Add(_Session.ReturnSubState)
			numReturns += 1
		End Sub

		Public Sub RemoveFile(ByVal _Session As SessionState)
			_Session.FileText = ReturnFileText(numReturns)
			_Session.StrokeTauntVal = ReturnStrokeTauntVal(numReturns)
			_Session.ReturnSubState = ReturnState(numReturns)
			ReturnFileText.RemoveAt(numReturns)
			ReturnStrokeTauntVal.RemoveAt(numReturns)
			ReturnState.RemoveAt(numReturns)
			numReturns -= 1
		End Sub

		Public Function getNumReturns() As Integer
			getNumReturns = numReturns
		End Function

		Public Sub resetList()
			numReturns = -1
			ReturnStrokeTauntVal = New List(Of Integer)
			ReturnFileText = New List(Of String)
			ReturnState = New List(Of String)
		End Sub

	End Class
End Class