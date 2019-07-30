# -*- coding: utf-8 -*-
"""
NameChange1.0

This is a program that automatically modifies
 the name of an word document.

 author:fanghao
"""
from docx import Document
import os

# 这个是放所有待修改的word文件的目录
dir_1 = "C:\\Users\\901\\Desktop\\123\\123"
filenames = os.listdir(dir_1)

# 自动修改
for a in range(len(filenames)):
    print(filenames[a])
    dir_docx = dir_1 + "\\" + filenames[a]
    try:
        document = Document(dir_docx)
    except:
        print("error")
    else:
        new_name = 'All_abstracts_' + str(a+752) + '.docx'
        try:
            os.rename(dir_1 + os.sep + filenames[a], dir_1 + os.sep + new_name)
        except(FileNotFoundError, FileExistsError, OSError):
            print("FileNotFoundError")

# vba语言
# 分页
# Option Explicit
#
# Sub SplitPagesAsDocuments()
#
#     Dim oSrcDoc As Document, oNewDoc As Document
#     Dim strSrcName As String, strNewName As String
#     Dim oRange As Range
#     Dim nIndex As Integer
#     Dim fso As Object
#
#     Set fso = CreateObject("Scripting.FileSystemObject")
#     Set oSrcDoc = ActiveDocument
#     Set oRange = oSrcDoc.Content
#
#     oRange.Collapse wdCollapseStart
#     oRange.Select
#
#     For nIndex = 1 To ActiveDocument.Content.Information(wdNumberOfPagesInDocument)
#         oSrcDoc.Bookmarks("\page").Range.Copy
#         oSrcDoc.Windows(1).Activate
#         Application.Browser.Target = wdBrowsePage
#         Application.Browser.Next
#
#         strSrcName = oSrcDoc.FullName
#         strNewName = fso.BuildPath(fso.GetParentFolderName(strSrcName), _
#                      fso.GetBaseName(strSrcName) & "_" & nIndex & "." & fso.GetExtensionName(strSrcName))
#         Set oNewDoc = Documents.Add
#         Selection.Paste
#         oNewDoc.SaveAs strNewName
#         oNewDoc.Close False
#     Next
#
#     Set oNewDoc = Nothing
#     Set oRange = Nothing
#     Set oSrcDoc = Nothing
#     Set fso = Nothing
#
#     MsgBox "结束！"
#
# End Sub

