'Desarrollado por Jorge M. Chávez
'Fecha: 01/03/2023

Sub Main
	IgnoreWarning(True)
	Call ReportReaderImport()		'D:\RUC1\DATA\Archivos fuente.ILB\2022_BCP_MIAMI.pdf
	Call AppendField()		'J_BCP_MIAMI2022b.IMD
	Call AppendField1()		'J_BCP_MIAMI2022b.IMD
	Call DirectExtraction()		'J_BCP_MIAMI2022b.IMD
	Call AppendField2()		'J_BCPMIAMI2022.IMD
	Call Summarization()		'J_BCPMIAMI2022.IMD
	Call ExportDatabaseXLSX()	'J_BCPMIAMI2022.IMD
	Client.CloseAll
	Client.DeleteDatabase "J_BCP_MIAMI2022b.IMD"
	Dim pm As Object
	Dim SourcePath As String
	Dim DestinationPath As String
	Set SourcePath = Client.WorkingDirectory
	Set DestinationPath = "D:\RUC1\DATA\_EECC"
	Client.RunAtServer False
	Set pm = Client.ProjectManagement
	pm.MoveDatabase SourcePath + "J_BCPMIAMI2022.IMD", DestinationPath
	pm.MoveDatabase SourcePath + "J.1_Resumen_BCPMIAMI.IMD", DestinationPath
	Set pm = Nothing
	Client.RefreshFileExplorer
End Sub


' Archivo - Asistente de importación: Report Reader
Function ReportReaderImport
	dbName = "J_BCP_MIAMI2022b.IMD"
	Client.ImportPrintReportEx "D:\RUC1\DATA\Definiciones de importación.ILB\BCPMIAMI_CTA_CTE.jpm", "D:\RUC1\DATA\Archivos fuente.ILB\2022_BCP_MIAMI.pdf", dbname, FALSE, FALSE
	Client.OpenDatabase (dbName)
End Function

' Anexar campo
Function AppendField
	Set db = Client.OpenDatabase("J_BCP_MIAMI2022b.IMD")
	Set task = db.TableManagement
	Set field = db.TableDef.NewField
	field.Name = "FECHA_PROC"
	field.Description = ""
	field.Type = WI_VIRT_CHAR
	field.Equation = "@split(@SpacesToOne(@remove(desc;""|""));""""; "" ""; 1; 0)"
	field.Length = 10
	task.AppendField field
	task.PerformTask
	Set task = Nothing
	Set db = Nothing
	Set field = Nothing
End Function

' Anexar campo
Function AppendField1
	Set db = Client.OpenDatabase("J_BCP_MIAMI2022b.IMD")
	Set task = db.TableManagement
	Set field = db.TableDef.NewField
	field.Name = "DESCRIPCION"
	field.Description = ""
	field.Type = WI_VIRT_CHAR
	field.Equation = "@split(@SpacesToOne(desc);fecha_proc; """"; 1; 0)"
	field.Length = 60
	task.AppendField field
	task.PerformTask
	Set task = Nothing
	Set db = Nothing
	Set field = Nothing
End Function


' Datos: Extracción directa
Function DirectExtraction
	Set db = Client.OpenDatabase("J_BCP_MIAMI2022b.IMD")
	Set task = db.Extraction
	task.AddFieldToInc "FECHA"
	task.AddFieldToInc "FECHA_PROC"
	task.AddFieldToInc "DESCRIPCION"
	task.AddFieldToInc "DEBITO"
	task.AddFieldToInc "CREDITO"
	task.AddFieldToInc "SALDO"
	dbName = "J_BCPMIAMI2022.IMD"
	task.AddExtraction dbName, "", ".NOT. (FECHA= """" .OR.  FECHA  = ""Debits"")"
	task.CreateVirtualDatabase = False
	task.PerformTask 1, db.Count
	Set task = Nothing
	Set db = Nothing
	Client.OpenDatabase (dbName)
End Function


' Anexar campo
Function AppendField2
	Set db = Client.OpenDatabase("J_BCPMIAMI2022.IMD")
	Set task = db.TableManagement
	Set field = db.TableDef.NewField
	field.Name = "PERIODO"
	field.Description = ""
	field.Type = WI_VIRT_CHAR
	field.Equation = """2022""+@Left(FECHA;2)"
	field.Length = 6
	task.AppendField field
	task.PerformTask
	Set task = Nothing
	Set db = Nothing
	Set field = Nothing
End Function

' Análisis: Resumen
Function Summarization
	Set db = Client.OpenDatabase("J_BCPMIAMI2022.IMD")
	Set task = db.Summarization
	task.AddFieldToSummarize "PERIODO"
	task.AddFieldToTotal "DEBITO"
	task.AddFieldToTotal "CREDITO"
	dbName = "J.1_Resumen_BCPMIAMI.IMD"
	task.OutputDBName = dbName
	task.CreatePercentField = FALSE
	task.StatisticsToInclude = SM_SUM
	task.PerformTask
	Set task = Nothing
	Set db = Nothing
	Client.OpenDatabase (dbName)
End Function

' Archivo-Exportar base de datos: XLSX
Function ExportDatabaseXLSX
	Set db = Client.OpenDatabase("J_BCPMIAMI2022.IMD")
	Set task = db.ExportDatabase
	task.IncludeAllFields
	eqn = ""
	task.PerformTask "D:\RUC1\DATA\Exportaciones.ILB\J_BCPMIAMI2022.XLSX", "Database", "XLSX", 1, db.Count, eqn
	Set db = Nothing
	Set task = Nothing
End Function