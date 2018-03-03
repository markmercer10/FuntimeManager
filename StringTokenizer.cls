VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "StringTokenizer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'local variable(s) to hold property value(s)
Private mvarSTtext As String 'local copy
Private mvarSTdelim As String 'local copy
Private mvarSTindex As Long 'local copy
Private mvarSThasMoreTokens As Boolean 'local copy
Public Sub Init(ByVal iniText As String, ByVal iniDelim As String)
    Let STtext = iniText
    Let STdelim = iniDelim
    Let STindex = 1
    Let SThasMoreTokens = True
End Sub

Public Function nextToken() As String
    Dim tempindex As Long
    tempindex = STindex
    If Not SThasMoreTokens Then
        MsgBox "Tokenizer has no more tokens"
        Exit Function
    End If
    
    Do
        If tempindex > Len(STtext) - 1 Then
            'nextToken = Trim(Mid$(STtext, STindex, tempindex - STindex))
            nextToken = Trim(MiD$(STtext, STindex))
            If Right$(nextToken, 1) = STdelim Then nextToken = MiD$(nextToken, 1, Len(nextToken) - 1)
            STindex = tempindex
            SThasMoreTokens = False
            Exit Do
        End If
        If InStr(1, STdelim, MiD$(STtext, tempindex, 1)) Then
            nextToken = MiD$(STtext, STindex, tempindex - STindex)
            STindex = tempindex + 1
            Exit Do
        End If
        tempindex = tempindex + 1
    Loop
End Function

Public Function hasMoreTokens() As Boolean
    Let hasMoreTokens = SThasMoreTokens
End Function

Public Property Let SThasMoreTokens(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.SThasMoreTokens = 5
    mvarSThasMoreTokens = vData
End Property


Public Property Get SThasMoreTokens() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.SThasMoreTokens
    SThasMoreTokens = mvarSThasMoreTokens
End Property



Public Property Let STindex(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.STindex = 5
    mvarSTindex = vData
End Property


Public Property Get STindex() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.STindex
    Let STindex = mvarSTindex
End Property



Public Property Let STdelim(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.STdelim = 5
    mvarSTdelim = vData
End Property


Public Property Get STdelim() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.STdelim
    STdelim = mvarSTdelim
End Property



Public Property Let STtext(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.STtext = 5
    mvarSTtext = vData
End Property


Public Property Get STtext() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.STtext
    STtext = mvarSTtext
End Property



