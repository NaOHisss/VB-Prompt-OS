Attribute VB_Name = "MathLib"
Public Type HashTable
Key As String
Value As String
End Type

Public Function getHashTable(str As String) As HashTable
Dim hash As HashTable
Dim tmp() As String
tmp = Split(str, "=")
hash.Key = tmp(0)
hash.Value = tmp(1)
getHashTable = hash
End Function
