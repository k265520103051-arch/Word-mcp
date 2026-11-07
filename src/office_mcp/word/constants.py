"""Word enums needed by the foundation. Tool modules keep their own enums."""

# WdSaveFormat: name -> (format value, file extension)
SAVE_FORMATS: dict[str, tuple[int, str]] = {
    "docx": (16, ".docx"),  # wdFormatXMLDocument
    "doc": (0, ".doc"),  # wdFormatDocument
    "pdf": (17, ".pdf"),  # wdFormatPDF
    "rtf": (6, ".rtf"),  # wdFormatRTF
    "txt": (2, ".txt"),  # wdFormatText
    "html": (8, ".html"),  # wdFormatHTML
    "odt": (23, ".odt"),  # wdFormatOpenDocumentText
}

# WdSaveOptions
wdDoNotSaveChanges = 0
wdSaveChanges = -1
