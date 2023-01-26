VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Simulacion 
   Caption         =   "Simulacion"
   ClientHeight    =   9360
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8325
   Icon            =   "Simulacion.dsx":0000
   OleObjectBlob   =   "Simulacion.dsx":1042A
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Simulacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CommandButton6_Click()
CommandButton4.Value = True
OptionButton1.Enabled = True
OptionButton1.Value = True
End Sub

Private Sub UserForm_Activate()
OptionButton1.Value = True
OptionButton2.Enabled = True
OptionButton3.Enabled = True
End Sub

Private Sub CommandButton1_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

Label8.Visible = False
Label9.Visible = False
Label5.Visible = True
Label6.Visible = True
End Sub

Private Sub CommandButton1_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

Label5.Visible = False
Label6.Visible = False
Label8.Visible = True
Label9.Visible = True
End Sub

Private Sub CommandButton2_Click()
Label8.Visible = True
MsgBox "Paso 2: Presionar Boton ROJO (Esto apagará la lampara Roja)."
End Sub


Private Sub CommandButton3_Click()

If OptionButton3.Enabled = True And OptionButton3.Visible = True Then
Label8.Visible = False
Label9.Visible = False
Label5.Visible = True
Label6.Visible = True
MsgBox "No puede regresar a condicion Normal, porque la falla continua."
Else
Label9.Visible = True
OptionButton1.Enabled = True
OptionButton1.Value = True
OptionButton3.Enabled = True
OptionButton3.Value = False
OptionButton2.Value = False
MsgBox "El sistema regresa a condicion Normal" & vbNewLine & vbNewLine & "(Esto sucede si las lamparas Verde y Roja se encuentran apagadas)"
End If
End Sub

Private Sub CommandButton5_Click()
Label7.Visible = False
Label4.Visible = True
CommandButton5.Visible = False
CommandButton4.Visible = True
MsgBox "El dispáro del sistema hacia el Relé RER620 no funcionará por haber desabilitado la llave."
End Sub

Private Sub CommandButton4_Click()
Label7.Visible = True
Label4.Visible = False
CommandButton5.Visible = True
CommandButton4.Visible = False
End Sub

Private Sub OptionButton1_Click()

CommandButton4.Visible = False
CommandButton5.Visible = True
Label7.Visible = True
Label8.Visible = True
Label9.Visible = True
Label4.Visible = False
Label5.Visible = False
Label6.Visible = False
OptionButton2.Enabled = True
OptionButton3.Enabled = True
OptionButton2.Value = False
OptionButton3.Value = False
End Sub

Private Sub OptionButton2_Click()
OptionButton1.Value = False
OptionButton1.Enabled = False
OptionButton3.Enabled = False
CommandButton1.Locked = True
Label8.Visible = False
Label9.Visible = False
Label5.Visible = True
Label6.Visible = True

MsgBox "Se encienden las lamparas Verde y Roja" & vbNewLine & vbNewLine & "Reposición a Condición Normal:" & vbNewLine & vbNewLine _
& "Paso 1: Presionar Boton VERDE. (Esto apagará la lampara Verde)"
End Sub

Private Sub OptionButton3_Click()
OptionButton1.Value = False
OptionButton2.Value = False
OptionButton1.Enabled = False
OptionButton2.Enabled = False
CommandButton1.Locked = True
Label8.Visible = False
Label9.Visible = False
Label5.Visible = True
Label6.Visible = True

MsgBox "Se encienden las lamparas Verde y Roja" & vbNewLine & vbNewLine & "Reposición a Condición Normal:" & vbNewLine & vbNewLine _
& "Paso 1: Presionar Boton VERDE. (Esto apagará la lampara Verde)"
End Sub


