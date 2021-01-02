#!user/Anaconda/envs/base python3
#-*-Coding:UTF8-*-
#VBA remake in python for commuting EXCEL files

import openpyxl

#https://stackoverflow.com/questions/9319317/quick-and-easy-file-dialog-in-python
from tkinter import filedialog
from tkinter import *
path = ""

#https://stackoverflow.com/questions/10377998/how-can-i-iterate-over-files-in-a-given-directory
import os
filename = ""





root = Tk()
root.withdraw()

path = filedialog.askdirectory()
print(path)

for filename in os.listdir(path):
    if filename.endswith(".xls") or filename.endswith(".xlsx"): 
        print(os.path.join(path, filename))
        continue
    else:
        continue




""" 

Sub TSSARE()


'
'                                   Declaration
'

Dim w As ThisWorkbook

Dim fajl As Variant
Dim fajlok As Variant
Dim id As Variant

Dim path As String

'
'                                   Folder Select
'

With Application.FileDialog(msoFileDialogFolderPicker)
    .Title = "Choose Folder!"

If .Show = True Then
path = .SelectedItems(1)


Else
path = ”nothing”

End If
End With

'
'                                   Datacopy
'

If path <> "nothing" Then

'
' i - Spacing for Macro Buttons, id-i = 1 becouse first id
'

i = 6
id = 1
'k = 0
Application.DisplayAlerts = False

fajl = Dir(path & "\*")

Do While fajl <> "" 'Or k < 10

Workbooks.Open (path & "\" & fajl)

ThisWorkbook.Sheets(1).Cells(i + 1, 1).Value = id
ThisWorkbook.Sheets(1).Cells(i + 1, 2).Value = Right(path, 8)
ThisWorkbook.Sheets(1).Cells(i + 1, 3).Value = fajl

'                                   Cover

ActiveWorkbook.Sheets(1).Cells(8, 4).Copy Destination:=ThisWorkbook.Sheets(1).Cells(i + 1, 4)
ActiveWorkbook.Sheets(1).Cells(9, 4).Copy Destination:=ThisWorkbook.Sheets(1).Cells(i + 1, 5)
ActiveWorkbook.Sheets(1).Cells(10, 4).Copy Destination:=ThisWorkbook.Sheets(1).Cells(i + 1, 6)
ActiveWorkbook.Sheets(1).Cells(11, 4).Copy Destination:=ThisWorkbook.Sheets(1).Cells(i + 1, 7)
ActiveWorkbook.Sheets(1).Cells(12, 4).Copy Destination:=ThisWorkbook.Sheets(1).Cells(i + 1, 8)
ActiveWorkbook.Sheets(1).Cells(13, 4).Copy Destination:=ThisWorkbook.Sheets(1).Cells(i + 1, 9)

'                                   Clipboard empty
Application.CutCopyMode = False



ActiveWorkbook.Close (False)

fajl = Dir()
i = i + 1
id = id + 1

'k = k + 1



Loop


End If

Application.DisplayAlerts = True

End Sub
"""




