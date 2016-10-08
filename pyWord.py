import win32com,os
from win32com.client import Dispatch, constants
#※※※※※※※※※※※※※※※※(变量定义)※※※※※※※※※※※※※※※※#
wordApp = win32com.client.Dispatch('Word.Application')
word = None             #当前处理的doc
Documents = None
ActiveDocument = None
Selection = None
#※※※※※※※※※※※※※※※※(方法定义)※※※※※※※※※※※※※※※※#
# 新建word
def newWord(fileaddr):
	global wordApp,word
	word = wordApp.Documents.Add()
	try:
		word.SaveAs(fileaddr)  #另存为
	except:
		print("word文件路径错误,无法保存.")

# 打开word
def openWord(fileaddr):
	global wordApp,word
	fileaddr = fileaddr.replace("\\","/")
	if(os.path.exists(fileaddr)):
		word = wordApp.Documents.Open(fileaddr)
		wordInit(True,False)
	else:
		print("文件打开失败,请检查路径是否正确,应改为/")

#初始化
def wordInit(visible=False,displayAlerts=False):
	global wordApp,Documents,ActiveDocument,Selection
	Documents = wordApp.Documents
	ActiveDocument = wordApp.ActiveDocument
	Selection = wordApp.Selection

# 后台运行,不显示,不警告
def setWordVisible(visible=True,displayAlerts=False):
	#运行情况可见(1为可见,0为不可见)
	wordApp.Visible = visible
	#屏蔽提示(1为显示,0为不显示)
	wordApp.DisplayAlerts = displayAlerts

#*******************************************************#
#                       页面设置                        #
#*******************************************************#
#设置页边距(默认单位为厘米)
def setPageMargin(top=2.54,bottom=2.54,width=0,height=0,left=3.18,right=3.18,header=1.5,footer=1.75):
	#设置页面方向,纵向=0,横向=1(与Excel不同)
	ActiveDocument.PageSetup.Orientation = 0
	#上边距(3cm,1cm=28.35pt)
	ActiveDocument.PageSetup.TopMargin = float(top)*28.35
	#下边距
	ActiveDocument.PageSetup.BottomMargin = float(bottom)*28.35
	#左边距
	ActiveDocument.PageSetup.LeftMargin  = float(left)*28.35
	#右边距
	ActiveDocument.PageSetup.RightMargin  = float(right)*28.35
	#页眉
	ActiveDocument.PageSetup.HeaderDistance = float(header)*28.35
	#页脚
	ActiveDocument.PageSetup.FooterDistance = float(footer)*28.35

#设置页眉
def setPageHeader(pos="center"):
	#定位到页眉文档
	wordApp.ActiveWindow.ActivePane.View.SeekView = 9
	#pos可选值(0为left,1为center,2为right)
	if(pos.lower() == "left"):
		wordApp.Selection.Paragraphs.Alignment = 0
	elif(pos.lower() == "center"):
		wordApp.Selection.Paragraphs.Alignment = 1
	elif(pos.lower() == "right"):
		wordApp.Selection.Paragraphs.Alignment = 2

	wordApp.Selection.Font.Name = "Calibri"
	wordApp.Selection.Font.Size = 10
	wordApp.Selection.Font.Bold = False
	wordApp.Selection.Font.Italic = False
	wordApp.Selection.TypeText("第")
	wordApp.Selection.Fields.Add(Selection.Range, -1, "PAGE", True)
	wordApp.Selection.TypeText("页,共")
	wordApp.Selection.Fields.Add(Selection.Range, -1, "NUMPAGES", True)
	wordApp.Selection.TypeText("页")
	#定位到主文档  
	wordApp.ActiveWindow.ActivePane.View.SeekView = 0      
	
#设置页脚
def setPageFooter(pos="center"):
	#定位到页脚文档
	wordApp.ActiveWindow.ActivePane.View.SeekView = 10
	#pos可选值(0为left,1为center,2为right)
	if(pos.lower() == "left"):
		wordApp.Selection.Paragraphs.Alignment = 0
	elif(pos.lower() == "center"):
		wordApp.Selection.Paragraphs.Alignment = 1
	elif(pos.lower() == "right"):
		wordApp.Selection.Paragraphs.Alignment = 2

	wordApp.Selection.Font.Name = "宋体"
	wordApp.Selection.Font.Size = 10
	wordApp.Selection.Font.Bold = False
	wordApp.Selection.Font.Italic = False
	wordApp.Selection.TypeText("第")
	wordApp.Selection.Fields.Add(Selection.Range, -1, "PAGE", True)
	wordApp.Selection.TypeText("页,共")
	wordApp.Selection.Fields.Add(Selection.Range, -1, "NUMPAGES", True)
	wordApp.Selection.TypeText("页")
	#定位到主文档  
	wordApp.ActiveWindow.ActivePane.View.SeekView = 0

#获取当前文档
def getActiveDocument():
	return ActiveDocument;

#*******************************************************#
#                       内容操作                        #
#*******************************************************#
#光标处插入内容
def insertContent(text,insertPos="after"):
	if(insertPos == "before"):
		Selection.InsertBefore(text)
		collapse(1)
	elif(insertPos == "after"):
		Selection.InsertAfter(text)
		collapse(0)
	#(HomeKey,EndKey)Unit=5为光标移到行首(行尾)或单元格文字之前(之后)
	# Selection.HomeKey(Unit=6)
	# s = ActiveDocument.Content.Start
	# e = ActiveDocument.Content.End
	# myRange = ActiveDocument.Range(s,e)
	

#删除内容
def deleteContent():
	Selection.Delete()

#全选
def allSelect():
	ActiveDocument.Content.Select()
	
#选择单个段落
def selectOneParagraphs(id=1):
	ActiveDocument.Paragraphs(str(id)).Range.Select()

#选择多个段落
def selectManyParagraphs(startId=1,endId=1):
	ActiveDocument.Range(ActiveDocument.Paragraphs(startId).Range.Start, ActiveDocument.Paragraphs(endId).Range.End).Select()

#选择Range
def selectRange(startPos=0,endPos=0):
	#获取光标的位置信息
	Selection.SetRange(startPos,endPos)

#返回段落数量
def getParagraphsCount():
	paragraphsCount = ActiveDocument.Paragraphs.Count
	return paragraphsCount

#返回图片数量
def getShapesCount():
	shapeCount = ActiveDocument.Shapes.Count
	# ActiveDocument.InlineShapes.Count
	return shapeCount


#*******************************************************#
#                       表格操作                        #
#*******************************************************#
#获取表格个数
def getTableCount():
	return ActiveDocument.Tables.Count

#插入表格
def insertTable(rowsNum=1,colsNum=1):
	#光标移到结尾
	Selection.EndKey(Unit=6)
	#插入换行
	# Selection.TypeParagraph()
	#插入表格
	ActiveDocument.Tables.Add(Range=Selection.Range,NumRows=rowsNum,NumColumns=colsNum)

#设置表格内容(tableDict为字典{'2*3':'abc'})
def setTableText(index=None,tableDict={}):
	if(type(index) == type(None)):
		return None
	table = ActiveDocument.Tables(index)
	for (key,value) in tableDict.items():
		rowNum = key.split("*")[0]
		colNum = key.split("*")[1]
		table.Cell(rowNum,colNum).Range.Text = value
		print(rowNum,colNum,value)

#获取表格内容
def getTableText(index=1):
	tableDict = {}
	table = ActiveDocument.Tables(index)
	rowsCount = table.Rows.Count
	colsCount = table.Columns.Count
	tableDict['rowsCount'] = rowsCount
	tableDict['colsCount'] = colsCount
	for rowNum in range(1,rowsCount+1):
		for colNum in range(1,colsCount+1):
			#异常用于处理合并单元格导致行列数不足
			try:
				text = replaceByte(table.Cell(rowNum,colNum).Range.Text)
				tableDict[str(rowNum)+"*"+str(colNum)] = text
			except:
				pass

	return tableDict

#获取表格坐标内容
def getTablePosText(index=1,rowNum=1,colNum=1):
	table = ActiveDocument.Tables(index)
	text = replaceByte(table.Cell(rowNum,colNum).Range.Text)
	return text

#表格增加一行
def tableAddRow(index=None):
	if(type(index) == type(None)):
		return None
	ActiveDocument.Tables(index).Rows.Add()


#表格增加一列
def tableAddCol(index=None):
	if(type(index) == type(None)):
		return None
	ActiveDocument.Tables(index).Columns.Add()
	#自动调整列宽
	ActiveDocument.Tables(index).AutoFitBehavior(2)


#*******************************************************#
#                       其他操作                        #
#*******************************************************#
#查找字符串(大小写，全部匹配)
def findContent(text,matchCase=True,matchWholeWord=True):
	Selection.Find.ClearFormatting()
	result = Selection.Find.Execute(FindText=str(text),MatchCase=matchCase,MatchWholeWord=matchWholeWord)
	Selection.Select()
	pageNum = Selection.Information(3)
	columnNum = Selection.Information(9)
	lineNum = Selection.Information(10)
	print(result,pageNum,lineNum,columnNum)
	return (result,pageNum,lineNum,columnNum)

#替换字符串
def replaceContent(oldText,newText):
	findObj = ActiveDocument.Content.Find
	findObj.ClearFormatting()
	findObj.Replacement.ClearFormatting()
	text = str(oldText)
	#是否区分大小写
	matchCase = True
	#是否全部匹配
	matchWholeWord = True
	#是否可使用查找通配符
	matchWildcards = False
	#是否查找同音字
	matchSoundsLike = False
	matchAllWordForms = False
	#向下搜索
	forward = True
	wrap = True
	#查找操作定位于格式或带格式的文本，而不是查找文本
	formats = True
	replaceWith = str(newText)
	#0为不替换，1为替换第一个，2为全部替换
	replace = 2
	result = findObj.Execute(text,matchCase,matchWholeWord,matchWildcards,matchSoundsLike,matchAllWordForms,forward,wrap,formats,replaceWith,replace)
	return result

#插入图片
def insertImage(imageaddr):
	try:
		if(os.path.exists(imageaddr)):
			Selection.InlineShapes.AddPicture(FileName=imageaddr, LinkToFile=0, SaveWithDocument=0)
		else:
			print("图片路径有误")
	except:
		print("图片插入失败")         

#移动光标
def moveCursor(direction="right",length=1,isSlect=0):
	if(direction == "right"):
		move = Selection.MoveRight(Unit=1,Count=length,Extend=isSlect)
		return move
	if(direction == "left"):
		move = Selection.MoveLeft(Unit=1,Count=length,Extend=isSlect)   
		return move

#将某一区域或所选内容折叠到起始位置或结束位置
def collapse(pos=0):
	#1为移动到开始位置，0为移动到结束位置
	Selection.Collapse(pos)

#忽略表格中特殊字符
def replaceByte(value):
	value = value.replace("\r","").replace("\x07","")
	return value

#打印预览
def printReview():
	ActiveDocument.PrintPreview()

#打印当前word
#printOption(0为正常打印，1为在一页打印正反面)
#pageSet为设置打印纸张(A3,A4)
#pageRange为打印范围(0为全部打印,2为打印当前页,3为设置起始结束页,4为设置打印范围)
#startPage为起始页码,endPage为结束页码
#Copies为打印份数
#Pages该参数表示要打印的页码和页码范围,以逗号分隔各项。例如,"2,6-10"表示打印第2页和第6至10页
#pageType(0为全部打印,1为打印奇数页,2为打印偶数页)
def printWord(printOption=0,pageSet="a4",pageRange=0,startPage=None,endPage=None,copies=1,pages="",pageType=0):
	if(printOption == 1):
		ActiveDocument.PageSetup.TwoPagesOnOne = True
	else:
		ActiveDocument.PageSetup.TwoPagesOnOne = False    
	if(pageSet.upper() == "A3"):
			#页面宽度
			ActiveDocument.PageSetup.PageWidth = 29.6*28.35
			#页面高度
			ActiveDocument.PageSetup.PageHeight = 41.91*28.35     
	elif(pageSet.upper() == "A4"):
			#页面宽度
			ActiveDocument.PageSetup.PageWidth = 21*28.35
			#页面高度
			ActiveDocument.PageSetup.PageHeight = 29.7*28.
	if(pages == ""):
		#获取总页数
		pages = Selection.Information(4)
		pages = "1-"+str(pages)
	ActiveDocument.PrintOut(False,False,pageRange,"",startPage,endPage,0,copies,pages,pageType)

#关闭word
def closeWord():
	#关闭前保存文档所做的更改
	Documents.Save()
	#关闭前显示提示
	wordApp.DisplayAlerts = 1
	if(type(word)!=type(None)):
		word.Close()
	
#退出word应用
def quitWord():
	if(type(wordApp)!=type(None)):
		wordApp.Quit()
		print("word程序执行完毕...")

