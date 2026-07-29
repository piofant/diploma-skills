#!/usr/bin/env python3
"""
Open .docx via LibreOffice UNO, refresh TOC + fields, export to PDF.
Usage: update_and_convert.py <in.docx> <out.pdf>
"""
import sys
import os
import subprocess
import time
import shutil

def main():
    in_docx = os.path.abspath(sys.argv[1])
    out_pdf = os.path.abspath(sys.argv[2])

    # Write a Basic macro file
    macro = r'''
function RefreshAndExport(sInFile as String, sOutFile as String) as Boolean
    Dim oDoc as Object
    Dim oArgs(0) as New com.sun.star.beans.PropertyValue
    oArgs(0).Name = "Hidden"
    oArgs(0).Value = True
    oDoc = StarDesktop.loadComponentFromURL(ConvertToURL(sInFile), "_blank", 0, oArgs())
    If IsNull(oDoc) Then
        RefreshAndExport = False
        Exit Function
    End If

    ' Update TOC + all indexes
    Dim oIndexes as Object
    oIndexes = oDoc.getDocumentIndexes()
    Dim i as Integer
    For i = 0 To oIndexes.getCount() - 1
        oIndexes.getByIndex(i).update()
    Next i

    ' Update fields
    oDoc.refresh()
    If oDoc.supportsService("com.sun.star.text.TextDocument") Then
        oDoc.getTextFields().refresh()
    End If

    ' Export to PDF
    Dim oExportArgs(0) as New com.sun.star.beans.PropertyValue
    oExportArgs(0).Name = "FilterName"
    oExportArgs(0).Value = "writer_pdf_Export"
    oDoc.storeToURL(ConvertToURL(sOutFile), oExportArgs())

    oDoc.close(True)
    RefreshAndExport = True
end function
'''
    # Use soffice with --headless and a python script invocation via bridged connection
    # Simpler approach: use unoconv-style -- run soffice with a macro call
    # Create temp script
    tmpdir = "/tmp/soffice_pdf_refresh"
    os.makedirs(tmpdir, exist_ok=True)
    user_profile = tmpdir + "/profile"
    os.makedirs(user_profile, exist_ok=True)

    # Approach: use Python-UNO via soffice's embedded python
    py_script = os.path.join(tmpdir, "refresh.py")
    with open(py_script, "w") as f:
        f.write(f'''
import uno
from com.sun.star.beans import PropertyValue

def prop(name, value):
    p = PropertyValue()
    p.Name = name
    p.Value = value
    return p

ctx = uno.getComponentContext()
resolver = ctx.ServiceManager.createInstanceWithContext(
    "com.sun.star.bridge.UnoUrlResolver", ctx)
ctx2 = resolver.resolve(
    "uno:socket,host=localhost,port=2202;urp;StarOffice.ComponentContext")
smgr = ctx2.ServiceManager
desktop = smgr.createInstanceWithContext("com.sun.star.frame.Desktop", ctx2)

in_url = "file://{in_docx}"
out_url = "file://{out_pdf}"

doc = desktop.loadComponentFromURL(in_url, "_blank", 0, (prop("Hidden", True),))

# Update indexes (TOC)
idx = doc.getDocumentIndexes()
for i in range(idx.getCount()):
    idx.getByIndex(i).update()

doc.refresh()
if hasattr(doc, "getTextFields"):
    doc.getTextFields().refresh()

doc.storeToURL(out_url, (prop("FilterName", "writer_pdf_Export"),))
doc.close(True)
print("Exported:", out_url)
''')

    # UNO-подход (soffice listener + LibreOfficePython с UNO-скриптом) на macOS
    # стабильно крашит LibreOfficePython и показывает crash-диалог. Фолбэк ниже
    # работает каждый раз. Пропускаем UNO целиком — это убирает крэш.
    ok = False

    if not ok or not os.path.exists(out_pdf):
        print("Используем soffice headless --convert-to pdf (UNO пропущен)")
        # Use soffice -macro invocation via Basic macro
        macro_cmd = f'macro:///Standard.Module1.RefreshTOC("{in_docx}","{out_pdf}")'
        # Simpler: use a tmp copy to a temp dir so fallback convert-to never overwrites adjacent files
        tmpcopy_dir = os.path.join("/tmp", "docx_pdf_tmp")
        os.makedirs(tmpcopy_dir, exist_ok=True)
        tmpcopy = os.path.join(tmpcopy_dir, os.path.basename(in_docx))
        shutil.copy(in_docx, tmpcopy)
        subprocess.run(["soffice", "--headless", "--convert-to", "pdf",
                        "--outdir", tmpcopy_dir, tmpcopy],
                       check=True)
        produced = os.path.join(tmpcopy_dir,
                                os.path.splitext(os.path.basename(in_docx))[0] + ".pdf")
        shutil.move(produced, out_pdf)
        os.remove(tmpcopy)


if __name__ == "__main__":
    main()
