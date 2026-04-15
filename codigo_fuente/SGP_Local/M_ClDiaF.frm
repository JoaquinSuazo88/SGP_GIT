VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{4DBFB8CD-9EF9-11D0-8BC4-00AA00B42B7C}#3.0#0"; "Cal32x30.ocx"
Begin VB.Form M_ClDiaF 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Calendario Días Feriados"
   ClientHeight    =   8085
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10275
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   8085
   ScaleWidth      =   10275
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   7455
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   9975
      Begin CalObjXLib.fpCalendar fpCalendar1 
         Height          =   6750
         Left            =   240
         TabIndex        =   1
         Top             =   480
         Width           =   9495
         _Version        =   196608
         _ExtentX        =   16748
         _ExtentY        =   11906
         _StockProps     =   68
         FirstDayOfWeek  =   1
         CurrentDate     =   "20000120"
         DateMin         =   "00000000"
         DateMax         =   "00000000"
         GrayAreaStyle   =   1
         GrayAreaBackColor=   -2147483633
         GrayAreaForeColor=   -2147483632
         HeaderStyle     =   2
         MonthHeaderStyle=   1
         YearHeaderStyle =   1
         BorderGrayAreaColor=   -2147483637
         ElementPictureBackground=   0   'False
         DisplayFormat   =   3
         BorderInnerStyle=   0
         BorderInnerHighlightColor=   -2147483633
         BorderInnerShadowColor=   -2147483642
         BorderInnerWidth=   1
         BorderOuterStyle=   0
         BorderOuterHighlightColor=   -2147483628
         BorderOuterShadowColor=   -2147483632
         BorderOuterWidth=   1
         BorderFrameWidth=   0
         BorderOutlineColor=   -2147483642
         BorderFrameColor=   -2147483633
         BorderOutlineWidth=   1
         BorderOutlineStyle=   1
         SpeedScrollYearIncrement=   1
         SpeedScrollMonthIncrement=   1
         GrayAreaAllowScroll=   0   'False
         WeekNumbers     =   0
         WeekDayHeader   =   3
         ElementTextStyle=   "M_ClDiaF.frx":0000
         DrawFocusRect   =   0
         MultiSelect     =   2
         YearStartMonth  =   1
         YearStartDay    =   1
         HeaderSeparatorWidth=   0
         HeaderSeparatorColor=   0
         YearFormatStyle =   2
         RangeBeginDate  =   "00000000"
         RangeEndDate    =   "00000000"
         GridLineColor   =   0
         GridLineStyle   =   3
         AutoSet         =   -1  'True
         InheritOverride =   1
         CompactFormat   =   ""
         MouseIcon       =   "M_ClDiaF.frx":02A9
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   10275
      _ExtentX        =   18124
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Style           =   1
      _Version        =   393216
      BorderStyle     =   1
   End
End
Attribute VB_Name = "M_ClDiaF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RS As New ADODB.Recordset
Dim ano As Long, mes As Long, dia As Long
Dim modo As String

Private Sub Form_Activate()
fg_descarga
End Sub

Private Sub Form_Load()
Me.HelpContextID = vg_OpcM
Me.Height = 8565
Me.Width = 10365
Msgtitulo = "Calendario Días Feriados"
fg_centra Me
modo = ""
Gl_Mo_Botones Me, 1
Gl_Ac_Botones Me, 1, 12, modo
MoverDiasFeriados
End Sub

Private Sub fpCalendar1_AfterSelection()
fpCalendar1.Element = ElementSelection
fpCalendar1.ElementIndex = Val(fg_pone_cero(ano, 4) & fg_pone_cero(mes, 2) & fg_pone_cero(dia, 2))
End Sub

Private Sub fpCalendar1_BeforeSelection(Cancel As Integer)
fpCalendar1.Element = ElementSelection
fpCalendar1.ElementIndex = Val(fg_pone_cero(ano, 4) & fg_pone_cero(mes, 2) & fg_pone_cero(dia, 2))
If fpCalendar1.ElementBackColor = &HFF& Then
   Cancel = True
   fpCalendar1.ElementBackColor = -2147483633 'colori
   fpCalendar1.ElementForeColor = vbBlack
Else
   Cancel = False
End If
End Sub

Private Sub fpCalendar1_DateChanging(Month As Integer, Day As Integer, Year As Integer, State As Integer, ByVal Shift As Integer, Cancel As Integer)
dia = Day
mes = Month
ano = Year
End Sub

Private Sub fpCalendar1_DblClick(CurrentMonth As Integer, CurrentDay As Integer, CurrentYear As Integer)
If Est Then Exit Sub
vg_nombre = ""
fpCalendar1.Element = ElementSpecificDate
fpCalendar1.ElementIndex = fg_pone_cero(CurrentYear, 4) & fg_pone_cero(CurrentMonth, 2) & fg_pone_cero(CurrentDay, 2)
'M_Feriado.LlenaDatos "", fg_pone_cero(CurrentDay, 2) & "/" & fg_pone_cero(CurrentMonth, 2) & "/" & fg_pone_cero(CurrentYear, 4), 2, Trim(fpCalendar1.ElementText)
'M_Feriado.Show 1, M_ClDiaF
'If Trim(vg_nombre) = "" Then Exit Sub
fpCalendar1.Element = ElementSpecificDate
fpCalendar1.ElementIndex = fg_pone_cero(CurrentYear, 4) & fg_pone_cero(CurrentMonth, 2) & fg_pone_cero(CurrentDay, 2)
fpCalendar1.ElementBackColor = &H8080FF      '&HFF&
fpCalendar1.ElementForeColor = vbBlack
fpCalendar1.ElementText = Trim(vg_nombre)
'fpCalendar1.MultiSelect = MultiSelectSimple
fpCalendar1.DrawFocusRect = CAL_DRAWFOCUSRECT_AROUND_TEXT
If modo = "" Then modo = "M"
Gl_Ac_Botones Me, 1, 0, modo: itab = 0
End Sub

Private Sub fpCalendar1_ViewChange(BeginMonth As Integer, BeginDay As Integer, BeginYear As Integer, EndMonth As Integer, EndDay As Integer, EndYear As Integer, Cancel As Integer)
Cancel = IIf(Toolbar1.Buttons(12).Visible = True, True, False)
If Cancel = False Then MoverDiasFeriados
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim fecdfe As Variant
Dim i As Long, j As Long
Select Case Button.Index
Case 1 '-------> Incluir
    modo = "A"
    vg_nombre = "": vg_codigo = ""
    M_ADiaFe.Show 1, M_Casino
    If Trim(vg_codigo) = "" Then Exit Sub
    PonerDiasFeriados vg_codigo, vg_nombre
    Gl_Ac_Botones Me, 1, 0, modo
Case 3 '-------> Alterar
Case 5 '-------> Borrar
    If MsgBox("Elimina registro...", vbQuestion + vbYesNo, Msgtitulo) = vbNo Then Exit Sub
    fpCalendar1.NextSelection = ""
    For i = 1 To fpCalendar1.SelCount
        fecdfe = fpCalendar1.NextSelection
        If Val(fecdfe) > 0 Then
           vg_dbpedweb.Execute "DELETE s_feriados WHERE ano = " & fg_pone_cero(Mid(fecdfe, 1, 4), 4) & " AND mes = " & fg_pone_cero(Mid(fecdfe, 5, 2), 2) & " AND dia = " & fg_pone_cero(Mid(fecdfe, 7, 2), 2) & ""
           fpCalendar1.Element = ElementSpecificDate
           fpCalendar1.ElementIndex = fecdfe
           fpCalendar1.ElementBackColor = -2147483633
           fpCalendar1.ElementText = ""
           fpCalendar1.ElementForeColor = vbBlack
        End If
     Next i
     modo = "": Gl_Ac_Botones Me, 1, 12, modo
Case 7 '-------> Actualizar lista
    MoverDiasFeriados
Case 10 '-------> Cancelar
    If MsgBox("Cancela...", vbQuestion + vbYesNo, Msgtitulo) = vbNo Then Exit Sub
    MoverDiasFeriados
    modo = "": Gl_Ac_Botones Me, 1, 12, modo
Case 12 '-------> Actualizar datos
    vg_dbpedweb.Execute "DELETE s_feriados WHERE ano = " & fpCalendar1.Year & ""
    For i = 1 To 12
        For j = 1 To Val(Mid(dEoM("01/" & fg_pone_cero(i, 2) & "/" & fpCalendar1.Year), 1, 2))
            fpCalendar1.Element = ElementSpecificDate
            fpCalendar1.ElementIndex = fpCalendar1.Year & fg_pone_cero(i, 2) & fg_pone_cero(j, 2)
            If fpCalendar1.ElementBackColor = &H8080FF Then
               vg_dbpedweb.Execute "INSERT INTO s_feriados (dia, mes, ano, Descripcion) VALUES (" & j & ", " & i & ", " & fpCalendar1.Year & ", '" & Trim(fpCalendar1.ElementText) & "')"
            End If
          Next j
    Next i
    modo = "": Gl_Ac_Botones Me, 1, 12, modo
Case 15 '------> Imprimir
    I_CalendarioDiasFeriados fpCalendar1.Year
Case 18 '-------> Cerrar
    Me.Hide
    Unload Me
End Select
End Sub

Private Sub MoverDiasFeriados()
Dim i As Long, j As Long
'fpCalendar1.CurrentDate = Now
fpCalendar1.Visible = False
fpCalendar1.AutoSet = True
fpCalendar1.DisplayFormat = 3
colori = fpCalendar1.ElementBackColor
colori = fpCalendar1.ElementBackColor
fpCalendar1.ShortDayName(1) = "Dom"
fpCalendar1.ShortDayName(2) = "Lun"
fpCalendar1.ShortDayName(3) = "Mar"
fpCalendar1.ShortDayName(4) = "Mie"
fpCalendar1.ShortDayName(5) = "Jue"
fpCalendar1.ShortDayName(6) = "Vie"
fpCalendar1.ShortDayName(7) = "Sab"
fpCalendar1.LongMonthName(1) = "Enero"
fpCalendar1.LongMonthName(2) = "Febrero"
fpCalendar1.LongMonthName(3) = "Marzo"
fpCalendar1.LongMonthName(4) = "Abril"
fpCalendar1.LongMonthName(5) = "Mayo"
fpCalendar1.LongMonthName(6) = "Junio"
fpCalendar1.LongMonthName(7) = "Julio"
fpCalendar1.LongMonthName(8) = "Agosto"
fpCalendar1.LongMonthName(9) = "Septiembre"
fpCalendar1.LongMonthName(10) = "Octubre"
fpCalendar1.LongMonthName(11) = "Noviembre"
fpCalendar1.LongMonthName(12) = "Diciembre"
For i = 1 To 12
    For j = 1 To Val(Mid(dEoM("01/" & fg_pone_cero(i, 2) & "/" & fpCalendar1.Year), 1, 2))
        fpCalendar1.Element = ElementSpecificDate
        fpCalendar1.ElementIndex = fpCalendar1.Year & fg_pone_cero(i, 2) & fg_pone_cero(j, 2)
        fpCalendar1.ElementBackColor = -2147483633
        fpCalendar1.ElementText = ""
        fpCalendar1.ElementForeColor = vbBlack
    Next j
Next i
      
Set RS = vg_dbpedweb.Execute("pedweb_s_diasferiados 1, " & fpCalendar1.Year & ", 0, 0")
Do While Not RS.EOF
    fpCalendar1.ElementIndex = RS!ano & fg_pone_cero(RS!mes, 2) & fg_pone_cero(RS!dia, 2)
    fpCalendar1.Element = ElementSpecificDate
    fpCalendar1.ElementIndex = RS!ano & fg_pone_cero(RS!mes, 2) & fg_pone_cero(RS!dia, 2)
    fpCalendar1.ElementText = IIf(IsNull(RS!descripcion), "", Trim(RS!descripcion))
    fpCalendar1.ElementBackColor = &H8080FF
    fpCalendar1.ElementForeColor = vbBlack
   RS.MoveNext
Loop
RS.Close: Set RS = Nothing
fpCalendar1.Visible = True
End Sub

Private Sub PonerDiasFeriados(op As String, dia As String)
Dim i As Long, j As Long
For i = 1 To 12
    For j = 1 To Val(Mid(dEoM("01/" & fg_pone_cero(i, 2) & "/" & fpCalendar1.Year), 1, 2))
        fpCalendar1.ElementIndex = fpCalendar1.Year & fg_pone_cero(i, 2) & fg_pone_cero(j, 2)
        If dia = "Ambos" And (fpCalendar1.ShortDayName(fg_Dia(fpCalendar1.Year & fg_pone_cero(i, 2) & fg_pone_cero(j, 2))) = "Sab" Or _
           fpCalendar1.ShortDayName(fg_Dia(fpCalendar1.Year & fg_pone_cero(i, 2) & fg_pone_cero(j, 2))) = "Dom") Then
           fpCalendar1.Element = ElementSpecificDate
           fpCalendar1.ElementIndex = fpCalendar1.Year & fg_pone_cero(i, 2) & fg_pone_cero(j, 2)
           fpCalendar1.ElementText = UCase(Mid(fg_Fecha_Dia(fpCalendar1.Year & fg_pone_cero(i, 2) & fg_pone_cero(j, 2), 1), 1, 3))
           fpCalendar1.ElementBackColor = IIf(op = "Incluir", &H8080FF, -2147483633)
           fpCalendar1.ElementForeColor = vbBlack
        ElseIf dia = "Domingo" And fpCalendar1.ShortDayName(fg_Dia(fpCalendar1.Year & fg_pone_cero(i, 2) & fg_pone_cero(j, 2))) = "Dom" Then
           fpCalendar1.Element = ElementSpecificDate
           fpCalendar1.ElementIndex = fpCalendar1.Year & fg_pone_cero(i, 2) & fg_pone_cero(j, 2)
           fpCalendar1.ElementText = UCase(Mid(fg_Fecha_Dia(fpCalendar1.Year & fg_pone_cero(i, 2) & fg_pone_cero(j, 2), 1), 1, 3))
           fpCalendar1.ElementBackColor = IIf(op = "Incluir", &H8080FF, -2147483633)
           fpCalendar1.ElementForeColor = vbBlack
        ElseIf dia = "Sabado" And fpCalendar1.ShortDayName(fg_Dia(fpCalendar1.Year & fg_pone_cero(i, 2) & fg_pone_cero(j, 2))) = "Sab" Then
           fpCalendar1.Element = ElementSpecificDate
           fpCalendar1.ElementIndex = fpCalendar1.Year & fg_pone_cero(i, 2) & fg_pone_cero(j, 2)
           fpCalendar1.ElementText = UCase(Mid(fg_Fecha_Dia(fpCalendar1.Year & fg_pone_cero(i, 2) & fg_pone_cero(j, 2), 1), 1, 3))
           fpCalendar1.ElementBackColor = IIf(op = "Incluir", &H8080FF, -2147483633)
           fpCalendar1.ElementForeColor = vbBlack
        End If
    Next j
Next i
End Sub
