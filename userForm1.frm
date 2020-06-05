VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "Month Selection"
   ClientHeight    =   6075
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10755
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub UserForm_Initialize()
    ComboBox1.List = Array("01", "02", "03", "04", "05", "06", "07", "08", "09", "10", "11", "12", "13", "14", "15", "16", "17", "18", "19", "20", "21", "22", "23", "24", "25", "26", "27", "28", "29", "30", "31")
    ComboBox2.List = Array("January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December")
    ComboBox3.List = Array("2010", "2011", "2012", "2013", "2014", "2015", "2016", "2017", "2018", "2019", "2020", "2021", "2022", "2023", "2024", "2025", "2026", "2027", "2028", "2029", "2030")
End Sub

Private Sub CommandButton1_Click()
    If Me.ComboBox1.Value = "" Or Me.ComboBox2.Value = "" Or Me.ComboBox3.Value = "" Then
        MsgBox "Please select the date before starting the script"
        Unload Me
    Else
        Me.Hide
        Call master(Me.ComboBox1.Value, Me.ComboBox2.Value, Me.ComboBox3.Value)
        Unload Me
    End If
End Sub

Private Sub CommandButton2_Click()
    Unload Me
End Sub
