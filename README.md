# xlsx_find_doc

帮研三学姐写的一个文件管理脚本,在某天上午,已经熬夜过的学姐诉求公司任务too shit,于是问我是否可以使用代码解决,一开始没有信息,没做过相关需求,要求下午就要,于是抱着电脑连肝了2大节课出来的

## 需求

有一个近千页的文档,一篇文章样式为粗体标题,下面为正文,后进行分页,后出现另一个文章
有一个excel上面有时间和文章标题
需要按照excel标题的顺序,按照excel给的标题在word中查找相应的正文,然后放到一个新的word按顺序排排放标题和正文,需要跟原来word一样格式的标题文字,需要分页

>由于时间比较急，代码没有那么pythonic

## 附加工具

wps专业版2016

## 思路

1. 将总的word拆分成一篇文章一个小文档(因为有些文章跨页数所以不能简单的一页一个文档)
2. 整理excel表,提取出所有的标题
3. 提取word文字,后面想到要将excel的标题与word的标题进行对比,但是word中文章和段落分隔不了,后来想到文章和标题有换行,然后就将文章拆分成段,二维列表第一行为第i篇文章,第j行为第i篇文章的第j段
4. 将excel标题和word段落进行对比
5. 若相同则提取出来并将该行全部打印到word中进行换行分页操作

## vba语言

在wps专业版下打开vba编辑器,源代码如下，可以使多个一个word文档拆分成一篇文章一个文档,

```vba
Option Explicit

Sub SplitPagesAsDocuments()

    Dim oSrcDoc As Document, oNewDoc As Document
    Dim strSrcName As String, strNewName As String
    Dim oRange As Range
    Dim nIndex As Integer
    Dim fso As Object

    Set fso = CreateObject("Scripting.FileSystemObject")
    Set oSrcDoc = ActiveDocument
    Set oRange = oSrcDoc.Content

    oRange.Collapse wdCollapseStart
    oRange.Select

    For nIndex = 1 To ActiveDocument.Content.Information(wdNumberOfPagesInDocument)
        oSrcDoc.Bookmarks("\page").Range.Copy
        oSrcDoc.Windows(1).Activate
        Application.Browser.Target = wdBrowsePage
        Application.Browser.Next

        strSrcName = oSrcDoc.FullName
        strNewName = fso.BuildPath(fso.GetParentFolderName(strSrcName), _
                     fso.GetBaseName(strSrcName) & "_" & nIndex & "." & fso.GetExtensionName(strSrcName))
        Set oNewDoc = Documents.Add
        Selection.Paste
        oNewDoc.SaveAs strNewName
        oNewDoc.Close False
    Next

    Set oNewDoc = Nothing
    Set oRange = Nothing
    Set oSrcDoc = Nothing
    Set fso = Nothing

    MsgBox "结束！"

End Sub
```

## tip

由于离开了工作站,自己笔记本cpu跟不上,在分页过程中出现vba编译器卡死或者闪退状态
