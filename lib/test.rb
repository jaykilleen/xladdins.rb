require 'win32ole'

path = 'C:\Users\KilleenJ\Projects\xladdins.rb\add-ins'

xl  = WIN32OLE.new('Excel.Application')
wb  = xl.Workbooks.Add()

xlmodule = wb.VBProject.VBComponents.Add(1)

example = 
'''
sub VBAMacro()
  msgbox "VBA Macro called"
end sub
'''

xlmodule.CodeModule.AddFromString(example)

# 55 references xlOpenXMLAddIn format
# This is not ideal as the id might change. I am unable to pass xlOpen
# https://msdn.microsoft.com/en-us/VBA/Excel-VBA/articles/xlfileformat-enumeration-excel
wb.SaveAs(path + '\example.xlam', 55)

xl.ActiveWorkbook.Close(0);
xl.Quit()