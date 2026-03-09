"""OOXML namespace constants for PPTX manipulation."""

NAMESPACE_A = "http://schemas.openxmlformats.org/drawingml/2006/main"
NAMESPACE_P = "http://schemas.openxmlformats.org/presentationml/2006/main"
NAMESPACE_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
NAMESPACE_RELS = "http://schemas.openxmlformats.org/package/2006/relationships"
NAMESPACE_P14 = "http://schemas.microsoft.com/office/powerpoint/2010/main"
NAMESPACE_A16 = "http://schemas.microsoft.com/office/drawing/2014/main"
NAMESPACE_CT = "http://schemas.openxmlformats.org/package/2006/content-types"
NAMESPACE_DCTERMS = "http://purl.org/dc/terms/"
NAMESPACE_XSI = "http://www.w3.org/2001/XMLSchema-instance"

NSMAP: dict[str, str] = {
    "a": NAMESPACE_A,
    "p": NAMESPACE_P,
    "r": NAMESPACE_R,
    "p14": NAMESPACE_P14,
    "a16": NAMESPACE_A16,
}

NSMAP_RELS: dict[str, str] = {
    "r": NAMESPACE_RELS,
}

NSMAP_CT: dict[str, str] = {
    "ct": NAMESPACE_CT,
}

REL_TYPE_SLIDE = (
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide"
)
REL_TYPE_NOTES_SLIDE = (
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide"
)
REL_TYPE_SLIDE_LAYOUT = (
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout"
)
REL_TYPE_NOTES_MASTER = (
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesMaster"
)
REL_TYPE_AUDIO = (
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/audio"
)
REL_TYPE_IMAGE = (
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"
)
REL_TYPE_MEDIA = "http://schemas.microsoft.com/office/2007/relationships/media"
REL_TYPE_THEME = (
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme"
)
