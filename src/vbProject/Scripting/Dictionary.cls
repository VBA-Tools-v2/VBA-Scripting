VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Dictionary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Drop-in replacement for Scripting.Dictionary on Mac"
''
' Dictionary v1.4.3
' (c) Tim Hall - https://github.com/timhall/VBA-Dictionary
'
' Native VBA implementation of `Scripting.Dictionary`
' https://docs.microsoft.com/en-us/office/vba/language/reference/user-interface-help/dictionary-object
'
' @author: tim.hall.engr@gmail.com
' @edited: andrew.pullon@radiuscore.co.nz
' @license: MIT (http://www.opensource.org/licenses/mit-license.php
' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ '
' RubberDuck Annotations
' https://rubberduckvba.com/ | https://github.com/rubberduck-vba/Rubberduck/
'
'@folder Scripting
'@moduledescription "Drop-in replacement for Scripting.Dictionary on Mac"
'@ignoremodule
'' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ '
Option Explicit

' --------------------------------------------- '
' Constants and Private Variables
' --------------------------------------------- '

#Const UseScriptingDictionaryIfAvailable = False

#If Mac Or Not UseScriptingDictionaryIfAvailable Then

' dict_KeyValue 0: FormattedKey, 1: OriginalKey, 2: Value
Private Type TDictionary
    KeyValues As Collection
    Keys() As Variant
    Items() As Variant
    ObjectKeys As Collection
    CompareMode As CompareMethod
End Type

#Else

Private Type TDictionary
    Dictionary As Object ' Scripting.Dictionary
End Type

#End If

Private This As TDictionary

' --------------------------------------------- '
' Types
' --------------------------------------------- '

Public Enum CompareMethod
    BinaryCompare = VBA.vbBinaryCompare
    TextCompare = VBA.vbTextCompare
    DatabaseCompare = VBA.vbDatabaseCompare
End Enum

' --------------------------------------------- '
' Properties
' --------------------------------------------- '

''
' Sets and returns the comparison mode for comparing string keys in a Dictionary object.
' https://docs.microsoft.com/en-us/office/vba/language/reference/user-interface-help/comparemode-property
'
' @property CompareMode
' @type {CompareMethod}
''
Public Property Get CompareMode() As CompareMethod
Attribute CompareMode.VB_Description = "Set or get the string comparison method."
#If Mac Or Not UseScriptingDictionaryIfAvailable Then
    CompareMode = This.CompareMode
#Else
    CompareMode = This.Dictionary.CompareMode
#End If
End Property
Public Property Let CompareMode(Value As CompareMethod)
#If Mac Or Not UseScriptingDictionaryIfAvailable Then
    If Me.Count > 0 Then
        ' Can't change CompareMode for Dictionary that contains data
        ' http://msdn.microsoft.com/en-us/library/office/gg278481(v=office.15).aspx
        Err.Raise 5 ' Invalid procedure call or argument
    End If

    This.CompareMode = Value
#Else
    This.Dictionary.CompareMode = Value
#End If
End Property

''
' Returns a Long (long integer) containing the number of items in a collection or Dictionary object. Read-only.
' https://docs.microsoft.com/en-us/office/vba/language/reference/user-interface-help/count-property-dictionary-object
'
' @property Count
' @type {Long}
''
Public Property Get Count() As Long
Attribute Count.VB_Description = "Get the number of items in the dictionary.\n"
#If Mac Or Not UseScriptingDictionaryIfAvailable Then
    Count = This.KeyValues.Count
#Else
    Count = This.Dictionary.Count
#End If
End Property

''
' Sets or returns an item for a specified key in a Dictionary object. Read/write.
' https://docs.microsoft.com/en-us/office/vba/language/reference/user-interface-help/item-property-dictionary-object
'
' @property Item
' @type {Variant}
' @param {Variant} Key | Key associated with the item being retrieved or added.
''
Public Property Get Item(Key As Variant) As Variant
Attribute Item.VB_Description = "Set or get the item for a given key."
Attribute Item.VB_UserMemId = 0
#If Mac Or Not UseScriptingDictionaryIfAvailable Then
    Dim dict_KeyValue As Variant
    dict_KeyValue = dict_GetKeyValue(Key)

    If Not IsEmpty(dict_KeyValue) Then
        If VBA.IsObject(dict_KeyValue(2)) Then
            Set Item = dict_KeyValue(2)
        Else
            Item = dict_KeyValue(2)
        End If
    Else
        ' Not found -> Returns Empty
    End If
#Else
    If VBA.IsObject(This.Dictionary.Item(Key)) Then
        Set Item = This.Dictionary.Item(Key)
    Else
        Item = This.Dictionary.Item(Key)
    End If
#End If
End Property
Public Property Let Item(Key As Variant, Value As Variant)
#If Mac Or Not UseScriptingDictionaryIfAvailable Then
    If Me.Exists(Key) Then
        dict_ReplaceKeyValue dict_GetKeyValue(Key), Key, Value
    Else
        dict_AddKeyValue Key, Value
    End If
#Else
    This.Dictionary.Item(Key) = Value
#End If
End Property
Public Property Set Item(Key As Variant, Value As Variant)
#If Mac Or Not UseScriptingDictionaryIfAvailable Then
    If Me.Exists(Key) Then
        dict_ReplaceKeyValue dict_GetKeyValue(Key), Key, Value
    Else
        dict_AddKeyValue Key, Value
    End If
#Else
    Set This.Dictionary.Item(Key) = Value
#End If
End Property

''
' Sets a key in a Dictionary object.
' https://docs.microsoft.com/en-us/office/vba/language/reference/user-interface-help/key-property
'
' @property Key
' @type {Variant}
' @param {Variant} Previous | The key value being changed.
' @param {Variant} Updated | New value that replaces the specified key.
''
Public Property Let Key(Previous As Variant, Updated As Variant)
Attribute Key.VB_Description = "Change a key to a different key."
#If Mac Or Not UseScriptingDictionaryIfAvailable Then
    Dim dict_KeyValue As Variant
    dict_KeyValue = dict_GetKeyValue(Previous)

    If Not VBA.IsEmpty(dict_KeyValue) Then
        dict_ReplaceKeyValue dict_KeyValue, Updated, dict_KeyValue(2)
    End If
#Else
    This.Dictionary.Key(Previous) = Updated
#End If
End Property

' ============================================= '
' Public Methods
' ============================================= '

''
' Adds a key and item pair to a Dictionary object.
' https://docs.microsoft.com/en-us/office/vba/language/reference/user-interface-help/add-method-dictionary
'
' @method Add
' @param {Variant} Key | The key associated with the item being added.
' @param {Variant} Item | The item associated with the key being added.
''
Public Sub Add(Key As Variant, Item As Variant)
Attribute Add.VB_Description = "Add a new key and item to the dictionary."
#If Mac Or Not UseScriptingDictionaryIfAvailable Then
    If Not Me.Exists(Key) Then
        dict_AddKeyValue Key, Item
    Else
        ' This key is already associated with an element of this collection
        Err.Raise 457
    End If
#Else
    This.Dictionary.Add Key, Item
#End If
End Sub

''
' Returns True if a specified key exists in the Dictionary object; False if it does not.
' https://docs.microsoft.com/en-us/office/vba/language/reference/user-interface-help/exists-method
'
' @method Exists
' @param {Variant} Key | Key value being searched for in the Dictionary object.
' @return {Boolean}
''
Public Function Exists(Key As Variant) As Boolean
Attribute Exists.VB_Description = "Determine if a given key is in the dictionary."
#If Mac Or Not UseScriptingDictionaryIfAvailable Then
    Exists = Not IsEmpty(dict_GetKeyValue(Key))
#Else
    Exists = This.Dictionary.Exists(Key)
#End If
End Function

''
' Returns an array containing all the items in a Dictionary object.
' https://docs.microsoft.com/en-us/office/vba/language/reference/user-interface-help/items-method
'
' @method Items
' @return {Array}
''
Public Function Items() As Variant
Attribute Items.VB_Description = "Get an array containing all items in the dictionary."
#If Mac Or Not UseScriptingDictionaryIfAvailable Then
    If Me.Count > 0 Then
        Items = This.Items
    Else
        Items = VBA.Split("") ' Creates initialized empty array that matches Dictionary Keys and Items.
    End If
#Else
    Items = This.Dictionary.Items
#End If
End Function

''
' Returns an array containing all existing keys in a Dictionary object.
' https://docs.microsoft.com/en-us/office/vba/language/reference/user-interface-help/keys-method
'
' @method Keys
' @return {Array}
''
Public Function Keys() As Variant
Attribute Keys.VB_Description = "Get an array containing all keys in the dictionary."
#If Mac Or Not UseScriptingDictionaryIfAvailable Then
    If Me.Count > 0 Then
        Keys = This.Keys
    Else
        Keys = VBA.Split("") ' Creates initialized empty array that matches Dictionary Keys and Items.
    End If
#Else
    Keys = This.Dictionary.Keys
#End If
End Function

''
' Removes a key/item pair from a Dictionary object.
' https://docs.microsoft.com/en-us/office/vba/language/reference/user-interface-help/remove-method-dictionary-object
'
' @method Remove
' @param {Variant} Key | Key associated with the key/item pair that you want to remove from the Dictionary object.
''
Public Sub Remove(Key As Variant)
Attribute Remove.VB_Description = "Remove a given key from the dictionary."
#If Mac Or Not UseScriptingDictionaryIfAvailable Then
    Dim dict_KeyValue As Variant
    dict_KeyValue = dict_GetKeyValue(Key)

    If Not VBA.IsEmpty(dict_KeyValue) Then
        dict_RemoveKeyValue dict_KeyValue
    Else
        Err.Raise 32811 ' Application-defined or object-defined error
    End If
#Else
    This.Dictionary.Remove Key
#End If
End Sub

''
' The RemoveAll method removes all key, item pairs from a Dictionary object.
' https://docs.microsoft.com/en-us/office/vba/language/reference/user-interface-help/removeall-method
'
' @method RemoveAll
''
Public Sub RemoveAll()
Attribute RemoveAll.VB_Description = "Remove all information from the dictionary."
#If Mac Or Not UseScriptingDictionaryIfAvailable Then
    Set This.KeyValues = New Collection

    Erase This.Keys
    Erase This.Items
#Else
    This.Dictionary.RemoveAll
#End If
End Sub

' ============================================= '
' Private Functions
' ============================================= '

#If Mac Or Not UseScriptingDictionaryIfAvailable Then

Private Function dict_GetKeyValue(dict_Key As Variant) As Variant
    On Error Resume Next
    dict_GetKeyValue = This.KeyValues(dict_GetFormattedKey(dict_Key))
    Err.Clear
End Function

Private Sub dict_AddKeyValue(dict_Key As Variant, dict_Value As Variant, Optional dict_Index As Long = -1)
    If Me.Count = 0 Then
        ReDim This.Keys(0 To 0)
        ReDim This.Items(0 To 0)
    Else
        ReDim Preserve This.Keys(0 To UBound(This.Keys) + 1)
        ReDim Preserve This.Items(0 To UBound(This.Items) + 1)
    End If

    Dim dict_FormattedKey As String
    dict_FormattedKey = dict_GetFormattedKey(dict_Key)

    If dict_Index >= 0 And dict_Index < This.KeyValues.Count Then
        ' Shift keys/items after + including index into empty last slot
        Dim dict_i As Long
        For dict_i = UBound(This.Keys) To dict_Index + 1 Step -1
            This.Keys(dict_i) = This.Keys(dict_i - 1)
            If VBA.IsObject(This.Items(dict_i - 1)) Then
                Set This.Items(dict_i) = This.Items(dict_i - 1)
            Else
                This.Items(dict_i) = This.Items(dict_i - 1)
            End If
        Next dict_i

        ' Add key/item at index
        This.Keys(dict_Index) = dict_Key
        If VBA.IsObject(dict_Value) Then
            Set This.Items(dict_Index) = dict_Value
        Else
            This.Items(dict_Index) = dict_Value
        End If

        ' Add key-value at proper index
        This.KeyValues.Add Array(dict_FormattedKey, dict_Key, dict_Value), dict_FormattedKey, Before:=dict_Index + 1
    Else
        ' Add key-value as last item
        If VBA.IsObject(dict_Key) Then
            Set This.Keys(UBound(This.Keys)) = dict_Key
        Else
            This.Keys(UBound(This.Keys)) = dict_Key
        End If
        If VBA.IsObject(dict_Value) Then
            Set This.Items(UBound(This.Items)) = dict_Value
        Else
            This.Items(UBound(This.Items)) = dict_Value
        End If

        This.KeyValues.Add Array(dict_FormattedKey, dict_Key, dict_Value), dict_FormattedKey
    End If
End Sub

Private Sub dict_ReplaceKeyValue(dict_KeyValue As Variant, dict_Key As Variant, dict_Value As Variant)
    Dim dict_Index As Long
    Dim dict_i As Integer

    dict_Index = dict_GetKeyIndex(dict_KeyValue(1))

    ' Remove existing dict_Value
    dict_RemoveKeyValue dict_KeyValue, dict_Index

    ' Add new dict_Key dict_Value back
    dict_AddKeyValue dict_Key, dict_Value, dict_Index
End Sub

Private Sub dict_RemoveKeyValue(dict_KeyValue As Variant, Optional ByVal dict_Index As Long = -1)
    Dim dict_i As Long
    If dict_Index = -1 Then
        dict_Index = dict_GetKeyIndex(dict_KeyValue(1))
    End If

    If dict_Index >= 0 And dict_Index <= UBound(This.Keys) Then
        ' Shift keys/items after index down
        For dict_i = dict_Index To UBound(This.Keys) - 1
            This.Keys(dict_i) = This.Keys(dict_i + 1)

            If VBA.IsObject(This.Items(dict_i + 1)) Then
                Set This.Items(dict_i) = This.Items(dict_i + 1)
            Else
                This.Items(dict_i) = This.Items(dict_i + 1)
            End If
        Next dict_i

        ' Resize keys/items to remove empty slot
        If UBound(This.Keys) = 0 Then
            Erase This.Keys
            Erase This.Items
        Else
            ReDim Preserve This.Keys(0 To UBound(This.Keys) - 1)
            ReDim Preserve This.Items(0 To UBound(This.Items) - 1)
        End If
    End If

    This.KeyValues.Remove dict_KeyValue(0)
    dict_RemoveObjectKey dict_KeyValue(1)
End Sub

Private Function dict_GetFormattedKey(dict_Key As Variant) As String
    If VBA.IsObject(dict_Key) Then
        dict_GetFormattedKey = dict_GetObjectKey(dict_Key)
    ElseIf VarType(dict_Key) = VBA.vbBoolean Then
        dict_GetFormattedKey = IIf(dict_Key, "-1__-1", "0__0")
    ElseIf VarType(dict_Key) = VBA.vbString Then
        dict_GetFormattedKey = dict_Key

        If Me.CompareMode = CompareMethod.BinaryCompare Then
            ' Collection does not have method of setting key comparison
            ' So case-sensitive keys aren't supported by default
            ' -> Approach: Append lowercase characters to original key
            '    AbC -> AbC___b_, abc -> abc__abc, ABC -> ABC_____
            Dim dict_Lowercase As String
            dict_Lowercase = vbNullString

            Dim dict_i As Integer
            Dim dict_Char As String
            Dim dict_Ascii As Integer
            For dict_i = 1 To VBA.Len(dict_GetFormattedKey)
                dict_Char = VBA.Mid$(dict_GetFormattedKey, dict_i, 1)
                dict_Ascii = VBA.Asc(dict_Char)
                If dict_Ascii >= 97 And dict_Ascii <= 122 Then
                    dict_Lowercase = dict_Lowercase & dict_Char
                Else
                    dict_Lowercase = dict_Lowercase & "_"
                End If
            Next dict_i

            If Not dict_Lowercase = vbNullString Then
                dict_GetFormattedKey = dict_GetFormattedKey & "__" & dict_Lowercase
            End If
        End If
    Else
        ' For numbers, add duplicate to distinguish from strings
        ' ->  123  -> "123__123"
        '    "123" -> "123"
        dict_GetFormattedKey = VBA.CStr(dict_Key) & "__" & CStr(dict_Key)
    End If
End Function

Private Function dict_GetObjectKey(dict_ObjKey As Variant) As String
    Dim dict_i As Integer
    For dict_i = 1 To This.ObjectKeys.Count
        If This.ObjectKeys.Item(dict_i) Is dict_ObjKey Then
            dict_GetObjectKey = "__object__" & dict_i
            Exit Function
        End If
    Next dict_i

    This.ObjectKeys.Add dict_ObjKey
    dict_GetObjectKey = "__object__" & This.ObjectKeys.Count
End Function

Private Sub dict_RemoveObjectKey(dict_ObjKey As Variant)
    Dim dict_i As Integer
    For dict_i = 1 To This.ObjectKeys.Count
        If This.ObjectKeys.Item(dict_i) Is dict_ObjKey Then
            This.ObjectKeys.Remove dict_i
            Exit Sub
        End If
    Next dict_i
End Sub

Private Function dict_GetKeyIndex(dict_Key As Variant) As Long
    Dim dict_i As Long
    For dict_i = 0 To UBound(This.Keys)
        If VBA.IsObject(This.Keys(dict_i)) And VBA.IsObject(dict_Key) Then
            If This.Keys(dict_i) Is dict_Key Then
                dict_GetKeyIndex = dict_i
                Exit For
            End If
        ElseIf VBA.IsObject(This.Keys(dict_i)) Or VBA.IsObject(dict_Key) Then
            ' Both need to be objects to check equality, skip
        ElseIf This.Keys(dict_i) = dict_Key Then
            dict_GetKeyIndex = dict_i
            Exit For
        End If
    Next dict_i
End Function

#End If

' ============================================= '
' Initialize & Terminate Methods
' ============================================= '

Private Sub Class_Initialize()
#If Mac Or Not UseScriptingDictionaryIfAvailable Then
    Set This.KeyValues = New Collection

    Erase This.Keys
    Erase This.Items
    Set This.ObjectKeys = New Collection
#Else
    Set This.Dictionary = CreateObject("Scripting.Dictionary")
#End If
End Sub

Private Sub Class_Terminate()
#If Mac Or Not UseScriptingDictionaryIfAvailable Then
    Set This.KeyValues = Nothing
    Set This.ObjectKeys = Nothing
#Else
    Set This.Dictionary = Nothing
#End If
End Sub
