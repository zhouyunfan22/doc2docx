from win32com import client as wc
import os
# wdFormatDocument = 0
# wdFormatDocument97 = 0
# wdFormatDocumentDefault = 16
# wdFormatDOSText = 4
# wdFormatDOSTextLineBreaks = 5
# wdFormatEncodedText = 7
# wdFormatFilteredHTML = 10
# wdFormatFlatXML = 19
# wdFormatFlatXMLMacroEnabled = 20
# wdFormatFlatXMLTemplate = 21
# wdFormatFlatXMLTemplateMacroEnabled = 22
# wdFormatHTML = 8
# wdFormatPDF = 17
# wdFormatRTF = 6
# wdFormatTemplate = 1
# wdFormatTemplate97 = 1
# wdFormatText = 2
# wdFormatTextLineBreaks = 3
# wdFormatUnicodeText = 7
# wdFormatWebArchive = 9
# wdFormatXML = 11
# wdFormatXMLDocument = 12
# wdFormatXMLDocumentMacroEnabled = 13
# wdFormatXMLTemplate = 14
# wdFormatXMLTemplateMacroEnabled = 15
# wdFormatXPS = 18
#单个文件转
def doc_to_docx_single(file):
    word = wc.Dispatch("Word.Application") # 打开word应用程序
    doc = word.Documents.Open(file) #打开word文件
    doc.SaveAs("{}x".format(file), 12)#另存为后缀为".docx"的文件，其中参数12指docx文件
    doc.Close() #关闭原来word文件
    word.Quit()
    print("完成！")
#遍历文件夹下的文件名称
def file_name(file_dir):
    for root, dirs, files in os.walk(file_dir):
        # 当前目录路径
        print(root)
        # 当前路径下所有子目录
        print(dirs)
        # 当前路径下所有非目录子文件
        print(files)
#批量转
def doc_to_docx_batch(file_dir,save_as_dir):
    word = wc.Dispatch("Word.Application")  # 打开word应用程序
    for root, dirs, files in os.walk(file_dir):
        # 当前路径下所有非目录子文件
        for f in files:
            print(f)
            doc = word.Documents.Open(file_dir+"/"+f)  # 打开word文件
            doc.SaveAs("{}x".format(save_as_dir+"/"+f), 12) #另存为后缀为".docx"的文件，其中参数12指docx文件
            doc.Close()  # 关闭原来word文件
            print("完成！")
    word.Quit()

if __name__ == '__main__':
    file_dir="C:/Users/"
    save_as_dir="C:/Users/"
    doc_to_docx_batch(file_dir,save_as_dir)


