# Launch LibreOffice with
#   /opt/libreoffice24.2/program/soffice --accept='socket,host=localhost,port=2024;urp;'

from scriptforge import ScriptForge, CreateScriptService
ScriptForge('localhost', 2024)

# Basic slide 8
basic = CreateScriptService('basic')


def main():
	# Execution context slide 9
	pf = CreateScriptService('platform')
	print(pf.OfficeVersion)
	fs = CreateScriptService('filesystem')
	print(fs.InstallFolder)
	fs.FileNaming = 'SYS'
	print(fs.InstallFolder)
	reg = CreateScriptService('region')
	print(reg.DayNames())
	print(reg.DayNames('de'))

	# Region service slide 10
	print(reg.Number2Text(73, 'fr'))
	print(reg.Number2Text(73, 'fr-BE'))

	# Exception, debugging, console
	exc = CreateScriptService('exception')
	exc.ConsoleClear()
	# exc.Console(modal = False)
	exc.DebugPrint('Exception', exc)
	exc.DebugPrint(reg.DayNames('fr'))

	# Calc gems (1) slide 13
	ui = CreateScriptService('ui')
	calc = ui.CreateDocument('calc')
	# calcrange = calc.OpenRangeSelector('Select a range of cells ...', closeafterselect = False)
	# print(calcrange)

	# Calc gems (2) slide 14
	calc.SetArray('B2', tuple(range(1, 11)))
	calc.SetArray('C2', tuple(range(2, 22, 2)))
	calc.SetArray('D2', tuple(range(3, 33, 3)))
	reg = calc.Region('B2')
	sumformula = '=SUM($%C1%R1:$%C2%R1)'
	formula = calc.Printf(sumformula, reg)		# =SUM($B2:$D2)
	formularange = calc.Offset(reg, columns = calc.Width(reg), width = 1)		# $Sheet1.$E$2:$E$11
	calc.SetFormula(formularange, formula)

	# Calc gems (3) slide 15
	reg = calc.Region('B2')		# $Sheet1.$B$2:$E$11
	calc.CurrentSelection = calc.CompactUp(reg, wholerow = True, filterformula = '=MOD($E2;10)=0')
	calc.RunCommand('Bold')

	# Menus slide 16
	menu = calc.CreateMenu('My Menu')
	menu.AddItem('About LibreOffice', command = 'About')
	calc.CloseDocument(False)

	# Extract data from the database slide 17
	db = CreateScriptService('Database', '', 'Bibliography')
	sql = """SELECT [Custom1], [Author], [Title] FROM [biblio] WHERE [Author] IS NOT NULL
	ORDER BY [Custom1], [Author], [Title]"""
	data = db.GetRows(sql)
	db.CloseDatabase()
	print(data)

	# Tree control in dialog  slide 18
	dialog = CreateScriptService('newdialog', 'mydialog', place = (10, 20, 150, 150))
	tree = dialog.CreateTreeControl('tree', place = (10, 10, dialog.Width - 5, dialog.Height - 5))
	root = tree.CreateRoot('Bibliography')
	tree.AddSubTree(root, data)
	dialog.Execute(modal = True)
	selection = tree.CurrentNode
	print(selection.DisplayValue)

	return


def makechart():
	# Slides 19-23
	ui = CreateScriptService('ui')
	fs = CreateScriptService('FileSystem')

	# Get Data
	db = CreateScriptService('database', registrationname = 'Bibliography')
	sql = 'SELECT [Custom1] AS [Language], [Identifier] FROM [biblio] ORDER BY [Language] ASC'
	data = db.GetRows(sql, header = True)
	db.CloseDatabase()

	# Import data in Calc
	calc = ui.CreateDocument('Calc', hidden = False)
	datarange = calc.SetArray('Sheet1.A1', data)
	pivot = calc.CreatePivotTable('Pivot1', datarange, targetcell = 'D1', datafields = 'Identifier;Count',
									rowfields = 'Language', rowtotals = False, columntotals = False,
									filterbutton = False)

	# Make and export chart
	chart = calc.CreateChart('NumberByLanguage', 'Sheet1', pivot, rowheader = True,
							 columnheader = True)
	chart.ChartType, chart.Dim3D, chart.Legend = 'Donut', True, True
	suffix = 'png'
	file = fs.GetTempName(suffix)
	chart.ExportToFile(file, suffix)
	# calc.CloseDocument(False)

	# Display chart in dialog (or form ...)
	dialog = CreateScriptService('NewDialog', place = (20, 20, 300, 300))
	picture = dialog.CreateImageControl('Picture', place = (5, 5, dialog.Width - 10, dialog.Height - 10),
										scale = 'KEEPRATIO')
	picture.Picture = file
	dialog.Execute()
	dialog.Terminate()


if __name__ == '__main__':
	choice = basic.InputBox('1 = Main, 2 = Chart', 'Ready for the show ?', '1')
	f = main if choice == '1' else makechart
	f()
