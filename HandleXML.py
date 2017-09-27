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

# 使用minidom解析器打开 XML 文档
aDocument = parse("plugin.xml")

assert isinstance(aDocument, Document)

rootElement = aDocument.documentElement

assert isinstance(rootElement, Element)

elementClassesList = rootElement.getElementsByTagName("Classes")


assert isinstance(elementClassesList, NodeList)

if elementClassesList.length > 1:
    print("Error:: There are more than one Classes Tag.")
    exit(1)

elementClasses = elementClassesList[0]

assert isinstance(elementClasses, Element)

elementItems = elementClasses.getElementsByTagName("Item")
