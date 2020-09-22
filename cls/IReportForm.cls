VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IReportForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Interface to implement polymorphism for report type forms."
Option Explicit

' ... Interface to general report forms
' ... such as those for APIs, COnstants, Types and Enums

Public Property Get ItemCount() As Long
'
End Property

Public Sub Init(ByRef pVBPInfo As VBPInfo, Optional ByRef pOK As Boolean = False, Optional ByRef pErrMsg As String = vbNullString)
'
End Sub

Public Sub ZOrder(Optional pOrder As Long = 0)
'
End Sub

