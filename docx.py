import docx

def read_docx(file_name):
    doc = docx.Document
    content = '\n'.join([para.text for para in doc.paragraphs])
    return content

# file_name = "C:\Samcef\Caesam\StrenBox_V2.4\customer\workspace\com.samcef.project.utilities\help\Chinese\HelpCE.html"
file_name = "D:\TMP_SC\Script\Help.docx"
con = read_docx(file_name)
print(con)