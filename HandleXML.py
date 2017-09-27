# -*- coding: utf-8 -*-
"""
Created on Fri Sep 22 14:52:52 2017
Purpose: Handle XML file with python.
@author: Jianguo (wjg155@163.com)
"""

from xml.dom.minidom import parse

from xml.dom.minidom import Document
from xml.dom.minidom import Element

from xml.dom.minicompat import NodeList
from test.test_compare import Cmp

from os import path
import os


def getHelpClassName(filePaht):
    # 使用minidom解析器打开 XML 文档
    aDocument = parse(filePaht)
    
    assert isinstance(aDocument, Document)
    
    rootElement = aDocument.documentElement
    
    assert isinstance(rootElement, Element)
    
    elementClassesList = rootElement.getElementsByTagName("Classes")
    
    
    assert isinstance(elementClassesList, NodeList)
    
    if elementClassesList.length > 1:
        print("Error:: There are more than one Classes Tag.")
        exit(1)
    
    if elementClassesList.length >= 1:
    
        elementClasses = elementClassesList[0]

        assert isinstance(elementClasses, Element)
        
        items = elementClasses.childNodes
        
        for item in items:
            if item.nodeType == 1:
                assert isinstance(item, Element)

                elementClass = item.getElementsByTagName("Class")[0]

                name = elementClass.childNodes[0].nodeValue
                standardname = "EO_Documentation"
                
                if standardname in name:
                    print(name)
                    elementResources = item.getElementsByTagName("Resources")[0]
                    elementUrl = elementResources.getElementsByTagName("Item")[0]

                    url = elementUrl.childNodes[0].nodeValue
                    print(url)    

def process(filePaht):
    # 使用minidom解析器打开 XML 文档
    aDocument = parse(filePaht[0])
    
    assert isinstance(aDocument, Document)
    
    rootElement = aDocument.documentElement
    
    assert isinstance(rootElement, Element)
    
    elementClassesList = rootElement.getElementsByTagName("Classes")
    
    
    assert isinstance(elementClassesList, NodeList)
    
    if elementClassesList.length > 1:
        print("Error:: There are more than one Classes Tag.")
        exit(1)
    
    if elementClassesList.length >= 1:
    
        elementClasses = elementClassesList[0]

        assert isinstance(elementClasses, Element)
        
        items = elementClasses.childNodes
        
        for item in items:
            if item.nodeType == 1:
                assert isinstance(item, Element)

                elementClass = item.getElementsByTagName("Class")[0]

                name = elementClass.childNodes[0].nodeValue
                standardname = "EO_Documentation"
                
                if standardname in name:
                    print(name)
                    elementResources = item.getElementsByTagName("Resources")[0]
                    elementUrl = elementResources.getElementsByTagName("Item")[0]

                    url = elementUrl.childNodes[0].nodeValue
                    
                    if "csm:documentation" in url:
                        url = "C:\Samcef\Caesam\StrenBox_V2.4\customer\workspace\com.samcef.project.utilities\help\Chinese" + "\\" + url[18:]
                    else:
                        url = "C:\Samcef\Caesam\StrenBox_V2.4\customer\workspace" + "\\" + filePaht[1] + "\help" + "\\" + os.path.split(url)[1]
                        
                    pdfurl = "C:\Samcef\Caesam\StrenBox_V2.4\customer\workspace" + "\\" + filePaht[1] + "\help" + "\\" + os.path.split(url)[1]
                    
                    pdfurl = pdfurl[0:(len(pdfurl)-4)] + "pdf"
                    
                    docurl = "C:\Samcef\Caesam\StrenBox_V2.4\customer\workspace" + "\\" + filePaht[1] + "\help" + "\\" + os.path.split(url)[1]
                    docurl = docurl[0:(len(docurl)-4)] + "docx"
                    
                    print(pdfurl)
                    print(docurl)
                    print(url)
                    
                    if os.path.isfile(url) == False:
                        print("Error::::::::::::::::::")
                    
                    savePDF(url, pdfurl, docurl)

def getPluginXMLList(wpath):
    pluginXMLList = []
    if path.exists(wpath):
        files = os.listdir(wpath)
        for file in files:
            if file != "com.samcef.common.utilities":
                ppath = os.path.join(wpath,file)
                if (os.path.isdir(ppath)):
                    for cpath in os.listdir(ppath):
                        cfile = os.path.join(ppath, cpath)
                        if os.path.isfile(cfile) and cpath == "plugin.xml":
                            pluginXMLList.append((cfile,file))
    return pluginXMLList

from win32com.client import Dispatch

def savePDF(htmlPath, pdfPath, wordPath):
    word = Dispatch('Word.Application')  # 打开word应用程序
    # word = DispatchEx('Word.Application') #启动独立的进程
    word.Visible = 0  # 后台运行,不显示
    word.DisplayAlerts = 0  # 不警告
    doc = word.Documents.Open(FileName=htmlPath, Encoding='gbk2312')
    
    doc.SaveAs(pdfPath, 17)
    doc.SaveAs(wordPath, 16)
    doc.Close()

# htmlPath = "C:\Samcef\Caesam\StrenBox_V2.4\customer\workspace\com.samcef.project.utilities\help\Chinese\HelpCE.html"
# pdfPath = pdfPath = "D:\TMP_SC\Script\HelpCE.pdf"
# savePDF(htmlPath, pdfPath)

wpath = "C:\Samcef\Caesam\StrenBox_V2.4\customer\workspace"
pluginXMLList = getPluginXMLList(wpath)
# for p in pluginXMLList:
#     print(p[0])
#     print(p[1])
for pluginXML in pluginXMLList:
#     print(pluginXML)
    process(pluginXML)

# filePath = "C:\Samcef\Caesam\StrenBox_V2.4\customer\workspace\com.samcef.static.basic.connectoranalysis\plugin.xml"
# getHelpClassName(filePath)
