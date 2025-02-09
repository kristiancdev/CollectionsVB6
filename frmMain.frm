VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "Form1"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    TestCollection
End Sub
Private Sub TestCollection()
    ' Crear una colección
    Dim col As New Collection
    
    ' Agregar elementos
    col.Add "Juan Pérez"
    col.Add "Ana Gómez"
    col.Add "Carlos López"
    
    ' Acceder a un elemento
    MsgBox col(2) ' Muestra "Ana Gómez"
    
    ' Recorrer la colección
    Dim item As Variant
    For Each item In col
        Debug.Print item
    Next item
    
    ' Eliminar un elemento
    col.Remove 2 ' Elimina el segundo elemento
    
    ' Contar elementos
    MsgBox "Número de elementos: " & col.Count
    
    ' Limpiar la colección
    Set col = New Collection
End Sub
