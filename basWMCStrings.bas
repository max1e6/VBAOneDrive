Attribute VB_Name = "basWMCStrings"
Option Explicit

'-----------------------------------------'
' Copyright William Max & Co. - June 2023 '
'-----------------------------------------'

Function gfHead(ByVal list As String, Optional delimiter As String = ",") As String

   '----------------------------------------------------------------------------------'
   ' Check if list has a trailing delimiter. If not then add delimiter to end of list '
   '----------------------------------------------------------------------------------'
   list = IIf(Right(list, Len(delimiter)) <> delimiter, list & delimiter, list)
   
   '---------------------------'
   ' Return first item in list '
   '---------------------------'
   gfHead = IIf(InStr(list, delimiter) - 1 < 1, "", Left$(list, InStr(list, delimiter) - 1))
   
End Function

Function gfTail(ByVal list As String, Optional delimiter As String = ",") As String
   '-------------------------'
   ' Return "tail" of string '
   '-------------------------'
   gfTail = IIf(InStr(list, delimiter) < 1, "", Mid$(list, InStr(list, delimiter) + Len(delimiter)))
End Function

Function gfReverseHead(ByVal list As String, Optional delimiter As String = ",") As String

   Dim elements As Variant
   
   elements = Split(list, delimiter)
   
   If UBound(elements) >= 0 Then
      gfReverseHead = elements(UBound(elements))
   Else
      gfReverseHead = ""
   End If

End Function

Function gfReverseTail(ByVal list As String, Optional delimiter As String = ",") As String
   
   Dim elements As Variant
   Dim iElement As Integer
   Dim sReturn As String
   
   elements = Split(list, delimiter)
   
   If UBound(elements) >= 0 Then
      sReturn = ""
      For iElement = 0 To UBound(elements) - 1
         sReturn = sReturn & elements(iElement) & delimiter
      Next
      gfReverseTail = Left(sReturn, Len(sReturn) - Len(delimiter))
   Else
      gfReverseTail = ""
   End If
End Function
