"""Shared XPath selectors used by PPTX XML helpers."""

XPATH_RELATIONSHIP_WITH_ID = "r:Relationship[@Id]"
XPATH_RELATIONSHIP_BY_ID = "r:Relationship[@Id='{rid}']"
XPATH_RELATIONSHIP_BY_TYPE = "r:Relationship[@Type='{rel_type}']"
XPATH_RELATIONSHIP_BY_TYPE_AND_TARGET = (
    "r:Relationship[@Type='{rel_type}'][@Target='{target}']"
)

XPATH_CT_DEFAULT_BY_EXTENSION = "ct:Default[@Extension='{extension}']"
XPATH_CT_OVERRIDE_BY_PATH_NAME = "ct:Override[@PartName='{path_name}']"

XPATH_NOTES_BODY_SHAPES = ".//p:ph[@type='body']/../../.."
XPATH_SHAPE_PARAGRAPHS = ".//a:p"
XPATH_PARAGRAPH_TEXT = ".//a:t"
XPATH_TXBODY_PARAGRAPHS = "a:p"
XPATH_NOTES_MASTER_ID_WITH_RID = ".//p:notesMasterId[@r:id]"

XPATH_P_CNVPR_WITH_ID = ".//p:cNvPr[@id]"
XPATH_P_SPTGT_WITH_SPID = ".//p:spTgt[@spid]"
XPATH_P_SPTGT_BY_SPID = ".//p:spTgt[@spid='{spid}']"
XPATH_P_CTN_WITH_ID = ".//p:cTn[@id]"
XPATH_TIMING_CONDS_WITH_DELAY = ".//p:par/p:cTn/p:stCondLst/p:cond[@delay]"

XPATH_P_PIC = ".//p:pic"
XPATH_PIC_CNVPR = "p:nvPicPr/p:cNvPr"
XPATH_PIC_AUDIO_FILE = "p:nvPicPr/p:nvPr/a:audioFile"
XPATH_PIC_MEDIA = "p:nvPicPr/p:nvPr/p:extLst/p:ext/p14:media"
XPATH_PIC_BLIP = "p:blipFill/a:blip"

XPATH_P_PAR = ".//p:par"
XPATH_P_AUDIO = ".//p:audio"
