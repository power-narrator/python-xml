"""Custom exceptions for PPTX operations."""


class PptxError(Exception):
    """Base exception for all PPTX-related errors."""

    pass


class InvalidPptxError(PptxError):
    """Raised when the PPTX file is malformed or invalid."""

    def __init__(self, path: str, reason: str) -> None:
        self.path = path
        self.reason = reason
        super().__init__(f"Invalid PPTX file '{path}': {reason}")


class SlideNotFoundError(PptxError):
    """Raised when a requested slide does not exist."""

    def __init__(self, slide_index: int, total_slides: int) -> None:
        self.slide_index = slide_index
        self.total_slides = total_slides
        super().__init__(
            f"Slide index {slide_index} out of range. "
            f"Presentation has {total_slides} slide(s)."
        )


class SlideXmlNotFoundError(PptxError):
    """Raised when a slide XML path cannot be found in workspace."""

    def __init__(self, slide_path: str) -> None:
        self.slide_path = slide_path
        super().__init__(f"Slide XML not found in workspace: '{slide_path}'")


class RelsNotFoundError(PptxError):
    """Raised when a .rels file does not exist in the PPTX archive."""

    def __init__(self, rels_path: str) -> None:
        self.rels_path = rels_path
        super().__init__(f"Relationships file not found in PPTX: '{rels_path}'")


class RelationshipIdNotFoundError(PptxError):
    """Raised when a required relationship ID is not found."""

    def __init__(self, source: str, rid: str) -> None:
        self.source = source
        self.rid = rid
        super().__init__(
            f"Relationship ID '{rid}' not found in relationships source '{source}'."
        )


class RelationshipTargetNotFoundError(PptxError):
    """Raised when a relationship target path resolves to a missing file."""

    def __init__(self, source: str, target: str) -> None:
        self.source = source
        self.target = target
        super().__init__(
            "Relationship target "
            f"'{target}' not found from relationships source '{source}'."
        )
