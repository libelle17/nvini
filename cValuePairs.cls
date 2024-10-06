VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cValuePairs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' The private collection used to hold the real data
Private m_cValuePairs As Collection

Private Sub Class_Initialize()
    ' explicit assignment is slightly faster than auto-instancing
    Set m_cValuePairs = New Collection
End Sub

' Add a new cValuePair item to the collection

Friend Sub Add(newItem As cValuePair, Optional Key As Variant)
Attribute Add.VB_Description = "Adds a member to a Collection object"
    ' TO DO: initialize new item's properties here
    ' ...
    ' add to the private collection
    m_cValuePairs.Add newItem, Key
End Sub

' Remove an item from the collection

Friend Sub Remove(index As Variant)
Attribute Remove.VB_Description = "Removes a member from a Collection object"
    m_cValuePairs.Remove index
End Sub

' Return a cValuePair item from the collection

Function Item(index As Variant) As cValuePair
Attribute Item.VB_Description = "Returns a specific member of a Collection object either by position or key"
Attribute Item.VB_UserMemId = 0
    Set Item = m_cValuePairs.Item(index)
End Function

' Return the number of items in the collection

Property Get Count() As Long
Attribute Count.VB_Description = "Returns the number of members in a collection"
    Count = m_cValuePairs.Count
End Property

' Remove all items from the collection

Friend Sub Clear()
Attribute Clear.VB_Description = "Removes all members from a Collection object"
    Set m_cValuePairs = New Collection
End Sub

' Implement support for enumeration (For Each)

Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    ' delegate to the private collection
    Set NewEnum = m_cValuePairs.[_NewEnum]
End Function

