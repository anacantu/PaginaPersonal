<%
  Class DynamicArray
  '************** Properties **************
  Private aData
  '****************************************

  '*********** Event Handlers *************
  Private Sub Class_Initialize()
    Redim aData(0)
  End Sub
  '****************************************

  '************ Property Get **************
  Public Property Get Data(iPos)
    'Make sure the end developer is not requesting an
    '"out of bounds" array element
    If iPos < LBound(aData) or iPos > UBound(aData) then
      Exit Property    'Invalid range
    End If

    Data = aData(iPos)
  End Property

  Public Property Get DataArray()
    DataArray = aData
  End Property
  '****************************************

  '************ Property Let **************
  Public Property Let Data(iPos, varValue)
    'Make sure iPos >= LBound(aData)
    If iPos < LBound(aData) Then Exit Property

    If iPos > UBound(aData) then
      'We need to resize the array
      Redim Preserve aData(iPos)
      aData(iPos) = varValue
    Else
      'We don't need to resize the array
      aData(iPos) = varValue
    End If
  End Property
  '****************************************


  '************** Methods *****************
  Public Function StartIndex()
     StartIndex = LBound(aData)
  End Function

  Public Function StopIndex()
     StopIndex = UBound(aData)
  End Function

  Public Sub Delete(iPos)
     'Make sure iPos is within acceptable ranges
     If iPos < LBound(aData) or iPos > UBound(aData) then
       Exit Sub    'Invalid range
     End If

     Dim iLoop
     For iLoop = iPos to UBound(aData) - 1
       aData(iLoop) = aData(iLoop + 1)
     Next

     Redim Preserve aData(UBound(aData) - 1)
  End Sub
  '****************************************
End Class
%>