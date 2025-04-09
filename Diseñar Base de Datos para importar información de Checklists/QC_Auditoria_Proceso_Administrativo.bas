Attribute VB_Name = "Módulo1"
'--- Definir hojas a utilizar
Dim HojaTapa As Range
Dim HojaDatos As Range
Dim HojaCiclo As Range
Dim HojaMetrica As Range

'-- Variables

'--- HOJA DATOS INICIALES
Public Proyecto As String
Public RF_Nombre As String
Public Empresa As String
Public Responsable As String
Public Revisor As String
Public Area As String
Public Nombre_Checklist As String
Public Version As String


'--- TABLA CHECKLIST TAMBIEN DATOS INICIALES

Public FechaInicio As String
Public FechaFin As String

'--- HOJA CICLO

Public Conceptos_Eval As String
Public Categoria As Integer

'--- HOJA METRICAS

Public Porcentaje_Falla_AUD As Double
Public Porcentaje_Eval As Double

'--- FUNCION CONEXION DB ---
Function ConexionDB() As Object
    Dim cn As Object
    Set cn = CreateObject("ADODB.Connection")
    cn.ConnectionString = "Driver=SQL Server;" & _
                         "Server=2035-CLSCL-102\QC_DOCUMENTAL;" & _
                         "Database=Checklist;" & _
                         "UID=sa;" & _
                         "PWD=Checklistqa!1"
    cn.Open
    Set ConexionDB = cn
End Function

'--- FUNCION TRANSFORMAR CATEGORIA A INT ---
Function ConvertirCategoria(estadoCategoria As String) As Integer
    Select Case Trim(LCase(estadoCategoria))
        Case "1.sin falla"
            ConvertirCategoria = 1
        Case "2.no aplica"
            ConvertirCategoria = 2
        Case "3.falla leve"
            ConvertirCategoria = 3
        Case "4.falla media"
            ConvertirCategoria = 4
        Case "5.falla grave"
            ConvertirCategoria = 5
        Case "6.falla invalidante"
            ConvertirCategoria = 6
        Case Else
            ConvertirCategoria = 0
    End Select
End Function

'--- VERIFICAR QUE LOS DATOS ESTAN INGRESADOS
Sub Verificar_Datos()

Set HojaTapa = ThisWorkbook.Worksheets("Tapa").Range("A1:B16")
Set HojaDatos = ThisWorkbook.Worksheets("Datos Iniciales").Range("A1:B12")

Proyecto = HojaDatos.Cells(4, 2)
Empresa = HojaDatos.Cells(5, 2)
Responsable = HojaDatos.Cells(6, 2)
Revisor = HojaDatos.Cells(7, 2)
FechaInicio = Format(HojaDatos.Cells(8, 2), "yyyymmdd")
FechaFin = Format(HojaDatos.Cells(9, 2), "yyyymmdd")
Area = HojaDatos.Cells(10, 2)
Version = HojaDatos.Cells(11, 2)

RF_Nombre = HojaTapa.Cells(10, 1)
Nombre_Checklist = HojaTapa.Cells(7, 1)

If Trim(Proyecto) = "" Then
    MsgBox "No se ha ingresado nombre Proyecto (Hoja Datos Iniciales)", vbInformation
ElseIf Trim(Empresa) = "" Then
    MsgBox "No se ha ingresado nombre Empresa (Hoja Datos Iniciales)", vbInformation
ElseIf Trim(Responsable) = "" Then
    MsgBox "No se ha ingresado nombre Responsable (Hoja Datos Iniciales)", vbInformation
ElseIf Trim(Revisor) = "" Then
    MsgBox "No se ha ingresado nombre Revisor (Hoja Datos Iniciales)", vbInformation
ElseIf Trim(FechaInicio) = "" Then
    MsgBox "No se ha ingresado Fecha de Inicio primer Ciclo (Hoja Datos Iniciales)", vbInformation
ElseIf Trim(FechaFin) = "" Then
    MsgBox "No se ha ingresado Fecha Fin Ciclo (Hoja Datos Iniciales)", vbInformation
ElseIf Trim(Area) = "" Then
    MsgBox "No se ha ingresado Módulo / Area / Opción (Hoja Datos Iniciales)", vbInformation
ElseIf Trim(RF_Nombre) = "" Or Trim(RF_Nombre) = "<<Nombre del requerimiento>>" Then
    MsgBox "No se ha ingresado Nombre del Requerimiento (Hoja Tapa)", vbInformation
ElseIf Trim(Version) = "" Or Trim(Version) = "Versión Documento <<X.X>>" Then
    MsgBox "No se ha ingresado Iteración del Documento (Hoja Datos Iniciales)", vbInformation
ElseIf Trim(Nombre_Checklist) = "" Then
    MsgBox "No se ha ingresado el nombre del Checklist (Hoja Tapa)", vbInformation
Else
MsgBox "Se han ingresado todos los datos correctamente"
End If

End Sub


Sub Guardar_Info()
 On Error GoTo ErrorHandler
 
 '--- MANEJAR CONEXIÓN Y SQL
 Dim cn As Object
 Dim rs As Object
Dim SQL As String


Dim FilaInicioAUD As Integer
Dim FilaFinAUD As Integer
'--- FILAS DE CONCEPTOS AUD
FilaInicioAUD = 4
FilaFinAUD = 9

'--- MANEJO ITERACIONES Y CATEGORIAS A INT
Dim i As Integer
Dim estadoCategoria As String

Set cn = ConexionDB()
If cn.State <> 1 Then
    MsgBox "No se pudo conectar a la base de datos", vbCritical
    Exit Sub
End If

Set HojaTapa = ThisWorkbook.Worksheets("Tapa").Range("A1:B16")
Set HojaDatos = ThisWorkbook.Worksheets("Datos Iniciales").Range("A1:B12")
Set HojaCiclo = ThisWorkbook.Worksheets("Ciclo 1").Range("A1:F55")
Set HojaMetrica = ThisWorkbook.Worksheets("Metricas").Range("A1:F25")

Proyecto = HojaDatos.Cells(4, 2)
Empresa = HojaDatos.Cells(5, 2)
Responsable = HojaDatos.Cells(6, 2)
Revisor = HojaDatos.Cells(7, 2)
FechaInicio = Format(HojaDatos.Cells(8, 2), "yyyymmdd")
FechaFin = Format(HojaDatos.Cells(9, 2), "yyyymmdd")
Area = HojaDatos.Cells(10, 2)
Version = HojaDatos.Cells(11, 2)

RF_Nombre = HojaTapa.Cells(10, 1)
Nombre_Checklist = HojaTapa.Cells(7, 1)

Porcentaje_Falla_AUD = HojaMetrica.Cells(2, 6)
Porcentaje_Eval = HojaMetrica.Cells(6, 6)

' --- VERIFICAR SI LOS DATOS YA EXISTEN EN LA TABLA DATOS
SQL = "SELECT id_datos FROM dbo.Datos WHERE Proyecto = '" & Replace(Proyecto, "'", "''") & "' " & _
"AND RF_req = '" & Replace(RF_Nombre, "'", "''") & "' " & _
"AND Nombre_res = '" & Replace(Responsable, "'", "''") & "' " & _
"AND Nombre_rev = '" & Replace(Revisor, "'", "''") & "' " & _
"AND Area = '" & Replace(Area, "'", "''") & "'"

    Set rs = cn.Execute(SQL)
    Dim iddatos As Long

If Not rs.EOF Then
'--- SI EXISTE, OBTENER EL ID_DATOS
iddatos = rs.Fields("id_datos").Value
' --- INSERTAR VALORES EN TABLA DATOS
Else
    SQL = "SET NOCOUNT ON; INSERT INTO dbo.Datos (Proyecto, RF_req, Nombre_res, Nombre_rev, Area) VALUES (" & _
    "'" & Replace(Proyecto, "'", "''") & "', " & _
    "'" & Replace(RF_Nombre, "'", "''") & "', " & _
    "'" & Replace(Responsable, "'", "''") & "', " & _
    "'" & Replace(Revisor, "'", "''") & "', " & _
    "'" & Replace(Area, "'", "''") & "'); SELECT SCOPE_IDENTITY() As NewID;"

    Set rs = cn.Execute(SQL)

    If rs Is Nothing Or rs.State = 0 Then
        MsgBox "Error: El recordset está cerrado o no se devolvió ningún resultado."
        Exit Sub
    End If

    If Not rs.EOF Then
        iddatos = rs.Fields(0).Value
    Else
        MsgBox "Error: No se devolvió SCOPE_IDENTITY() en la inserción de Datos."
        Exit Sub
    End If
End If

'--- INSERTAR VALORES EN TABLA METRICAS

    SQL = "SET NOCOUNT ON; INSERT INTO dbo.Metricas (porcentaje_falla_aud, porcentaje_eval, id_datos) VALUES (" & _
CStr(Replace(CStr(Porcentaje_Falla_AUD), ",", ".")) & ", " & _
CStr(Replace(CStr(Porcentaje_Eval), ",", ".")) & ", " & _
CStr(iddatos) & "); SELECT SCOPE_IDENTITY()"

Set rs = cn.Execute(SQL)
Dim idmetricas As Long

If rs Is Nothing Or rs.State = 0 Then
    MsgBox "Error: El recordset está cerrado o no se devolvió ningún resultado."
    Exit Sub
End If

If Not rs.EOF Then
    idmetricas = rs.Fields(0).Value
Else
    MsgBox "Error: No se devolvió SCOPE_IDENTITY() en la inserción de Datos."
    Exit Sub
End If


'--- INSERTAR VALORES EN TABLA CHECKLIST

    SQL = "INSERT INTO dbo.Checklist (nombre_checklist, version, fecha_inicio, fecha_fin, id_metricas) VALUES (" & _
"'" & Replace(Nombre_Checklist, "'", "''") & "', " & _
"'" & Replace(Version, "'", "''") & "', " & _
"'" & Replace(FechaInicio, "'", "''") & "', " & _
"'" & Replace(FechaFin, "'", "''") & "', " & _
CStr(idmetricas) & ")"

cn.Execute SQL

'--- INSERTAR VALORES EN TABLA AUD

For i = FilaInicioAUD To FilaFinAUD
        Conceptos_Eval = Trim(HojaCiclo.Cells(i, 1))
        estadoCategoria = Trim(LCase(HojaCiclo.Cells(i, 4).Value))
    
        Categoria = ConvertirCategoria(estadoCategoria)

    SQL = "INSERT INTO dbo.Evaluacion_AUD (conceptos_eval, id_valor, id_metricas) VALUES (" & _
"'" & Replace(Conceptos_Eval, "'", "''") & "', " & _
CStr(Categoria) & ", " & _
CStr(idmetricas) & ")"

cn.Execute SQL

    Next i



MsgBox "Se guardaron los datos en Base de datos QA_DOCUMENTAL. " & Format(Now, "yyyy.mm.dd HH:mm:ss")


CleanUp:
    If Not cn Is Nothing Then
        If cn.State = 1 Then cn.Close
        Set cn = Nothing
        Set HojaMetrica = Nothing
    End If
    Exit Sub
    
ErrorHandler:
    MsgBox "Error: " & Err.Description, vbCritical
    Resume CleanUp

End Sub

Sub Contraseña()

    Dim ContraseñaCorrecta As String
    ContraseñaCorrecta = "123"

    Dim ContraseñaIngresada As String
    ContraseñaIngresada = InputBox("Por favor, ingrese la contraseña para continuar:", "Contraseña Requerida")

    ' Validar la contraseña
    If ContraseñaIngresada <> ContraseñaCorrecta Then
        MsgBox "Contraseña incorrecta.", vbCritical
        Exit Sub
    End If

    MsgBox "Contraseña correcta", vbInformation

    Call Guardar_Info
End Sub
    



