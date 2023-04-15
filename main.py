import json
import win32com.client


def readProp(doc):
    try:
        for prop in doc.BuiltInDocumentProperties:
            print("0 : 1", prop.name, prop.value)
    except Exception as e:
        print('\n\n', e)


def setProp(doc):
    with open("jexmaple.json", encoding="utf-8") as json_file:
        data_json = json.load(json_file)
    try:
        for key_json in data_json:
            try:
                for prop in doc.BuiltInDocumentProperties:
                    if prop.name == key_json:
                        doc.BuiltInDocumentProperties[prop.name] = data_json[key_json]
                        break
                else:
                    print("Attribute not found")
            except Exception as e:
                print('\n\n', e)
    except Exception as e:
        print('\n\n', e)


if __name__ == '__main__':
    word = win32com.client.Dispatch("Word.Application")
    word.Visible = 0
    doc_path = "C:\\Users\\Andrey\\Downloads\\Doc2.docx"
    doc = word.Documents.Open(doc_path)
    doc.Saved = False
    ##READ ATTRIBUTE FROM PROP
    readProp(doc)
    ##SET ATRIBUTE FROM PROP
    setProp(doc)
    doc.Close()
    word.Quit()
