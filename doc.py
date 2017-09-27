# coding=utf-8
import win32com
from win32com.client import Dispatch, DispatchEx, CDispatch

word = Dispatch('Word.Application')  # 打开word应用程序
# word = DispatchEx('Word.Application') #启动独立的进程
word.Visible = 0  # 后台运行,不显示
word.DisplayAlerts = 0  # 不警告
path = 'D:\TMP_SC\Script\Help.docx'  # word文件路径
path = "C:\Samcef\Caesam\StrenBox_V2.4\customer\workspace\com.samcef.project.utilities\help\Chinese\HelpCE.html"
doc = word.Documents.Open(FileName=path, Encoding='gbk2312')
pdfPath = "D:\TMP_SC\Script\HelpCE.pdf"

doc.SaveAs(pdfPath, 17)

# content = doc.Range(doc.Content.Start, doc.Content.End)
# content = doc.Range()
print ('----------------')
print ('段落数: ' + str(doc.Paragraphs.count))

# 利用下标遍历段落
for i in range(len(doc.Paragraphs)):
    para = doc.Paragraphs[i]
    print (para.Range.text)
print ('-------------------------')

# 直接遍历段落
for para in doc.paragraphs:
    print (para.Range.text)
    # print para  #只能用于文档内容全英文的情况

doc.Close()  # 关闭word文档
# word.Quit  #关闭word程序