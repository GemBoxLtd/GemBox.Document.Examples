# Create ComHelper object and set license.
# NOTE: If you're using a Professional version you'll need to put your serial key below.
import win32com.client as COM
comHelper = COM.Dispatch("GemBox.Document.ComHelper")
comHelper.ComSetLicense("FREE-LIMITED-KEY")

# Load input document.
import os
document = comHelper.Load(os.getcwd() + "\\ComTemplate.docx")

# Find and replace.
document.Content.Replace("PLACEHOLDER1", "Sample Value 1")
document.Content.Replace("PLACEHOLDER2", "Sample Value 2")
document.Content.Replace("PLACEHOLDER3", "Sample Value 3")

# Mail merge.
source = COM.Dispatch("System.Collections.Hashtable")
source.Add("Name", "John")
source.Add("Surname", "Doe")
source.Add("Age", 30)
document.MailMerge.Execute(source)

# Modify bookmarks.
document.Bookmarks.Item("Bookmark1").GetContent(True).LoadText("Sample Content 1.")
document.Bookmarks.Item("Bookmark2").GetContent(True).LoadText("Sample Content 2.")

# Get output path and save document as PDF file.
path = os.getcwd()  + "\\ComExample.pdf"

document.Save(path)
print("Document saved as '" + path + "'")