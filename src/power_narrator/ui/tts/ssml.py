"""SSML processing for custom slide syntax."""

import re
from abc import ABC, abstractmethod
from re import Match
from xml.sax.saxutils import escape

from typing_extensions import override


class SSMLRule(ABC):
    """Base class for SSML transformation rules."""

    _pattern: re.Pattern[str]

    @classmethod
    @abstractmethod
    def _replacement(cls, match: Match[str]) -> str:
        """Generate replacement string for a regex match.

        Args:
            match: The regex match object.

        Returns:
            The replacement string.
        """
        pass

    @classmethod
    def apply(cls, text: str) -> str:
        """Apply the rule to the input text.

        Args:
            text: The input text to transform.

        Returns:
            The transformed text.
        """
        return cls._pattern.sub(cls._replacement, text)


class VoiceRule(SSMLRule):
    """Wrap paragraphs that start with [voice-name] in <voice> tags."""

    _pattern = re.compile(r"^\[([^\]]+)\](.*)$", re.MULTILINE)

    @override
    @classmethod
    def _replacement(cls, match: Match[str]) -> str:
        return f'<voice name="{match.group(1)}">{match.group(2)}</voice>'


class BreakRule(SSMLRule):
    """Convert tilde runs surrounded by spaces to <break> tags."""

    _pattern = re.compile(r"(?<!\S)~+(?!\S)")

    @override
    @classmethod
    def _replacement(cls, match: Match[str]) -> str:
        return f'<break time="{len(match.group(0)) * 0.5}s"/>'


class EmphasisRule(SSMLRule):
    """Convert _text_ to <emphasis level="strong">text</emphasis>."""

    _pattern = re.compile(r"(?<!\S)_(.+?)_(?!\S)")

    @override
    @classmethod
    def _replacement(cls, match: Match[str]) -> str:
        return f'<emphasis level="strong">{match.group(1)}</emphasis>'


class SSMLProcessor:
    """Apply a sequence of SSML transformation rules."""

    # Apply voice rule last as it removes new line character after content
    _rules: list[type[SSMLRule]] = [BreakRule, EmphasisRule, VoiceRule]

    @classmethod
    def to_ssml(cls, text: str) -> str:
        """Convert custom syntax to SSML wrapped in <speak> tags.

        Args:
            text: The input text with custom syntax.

        Returns:
            The SSML-formatted text.
        """
        escaped = escape(text, {'"': "&quot;"})

        for rule in cls._rules:
            escaped = rule.apply(escaped)

        return f"<speak>{escaped}</speak>"
