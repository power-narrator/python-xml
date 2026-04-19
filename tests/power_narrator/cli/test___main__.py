import json
from pathlib import Path

from power_narrator.cli.__main__ import main


class FakePptxFile:
    def __init__(self) -> None:
        self.notes = ["original"]
        self.exported_to: Path | None = None
        self.opened_path: Path | None = None

    @classmethod
    def open(cls, path: Path) -> "FakePptxFile":
        instance = cls()
        instance.opened_path = path
        return instance

    def __enter__(self) -> "FakePptxFile":
        return self

    def __exit__(self, exc_type, exc_value, traceback) -> None:
        return None

    def get_slides(self) -> list[dict[str, object]]:
        return [{"notes": note, "audio": []} for note in self.notes]

    def set_slide_notes(self, slide_index: int, notes: str) -> None:
        if slide_index >= len(self.notes):
            raise IndexError("slide out of range")

        self.notes[slide_index] = notes

    def save_notes(self) -> None:
        return None

    def save_audio_for_slide(self, slide_index: int, mp3_path: Path) -> None:
        if not isinstance(mp3_path, Path):
            raise TypeError("mp3_path must be Path")

        if slide_index >= len(self.notes):
            raise IndexError("slide out of range")

    def export_to(self, output_path: Path) -> None:
        self.exported_to = output_path


def test_cli_writes_results_and_exports_when_all_ops_succeed(
    tmp_path, monkeypatch
) -> None:
    fake = FakePptxFile()
    monkeypatch.setattr("power_narrator.cli.__main__.PptxFile.open", lambda path: fake)

    request_path = tmp_path / "request.json"
    results_path = tmp_path / "results.json"
    request_path.write_text(
        json.dumps(
            {
                "input": "input.pptx",
                "output": "output.pptx",
                "ops": [
                    {"op": "get_slides", "args": {}},
                    {
                        "op": "set_slide_notes",
                        "args": {"slide_index": 0, "notes": "updated"},
                    },
                    {
                        "op": "save_audio_for_slide",
                        "args": {"slide_index": 0, "mp3_path": "audio.mp3"},
                    },
                ],
            }
        ),
        encoding="utf-8",
    )

    exit_code = main([str(request_path), str(results_path)])
    results = json.loads(results_path.read_text(encoding="utf-8"))

    assert exit_code == 0
    assert results == {
        "results": [
            {
                "success": True,
                "result": [{"notes": "original", "audio": []}],
                "message": "",
            },
            {"success": True, "result": None, "message": ""},
            {"success": True, "result": None, "message": ""},
        ]
    }
    assert fake.notes == ["updated"]
    assert fake.exported_to == Path("output.pptx")


def test_cli_skips_export_when_any_operation_fails(tmp_path, monkeypatch) -> None:
    fake = FakePptxFile()
    monkeypatch.setattr("power_narrator.cli.__main__.PptxFile.open", lambda path: fake)

    request_path = tmp_path / "request.json"
    results_path = tmp_path / "results.json"
    request_path.write_text(
        json.dumps(
            {
                "input": "input.pptx",
                "output": "output.pptx",
                "ops": [
                    {
                        "op": "set_slide_notes",
                        "args": {"slide_index": 5, "notes": "updated"},
                    },
                    {"op": "get_slides", "args": {}},
                ],
            }
        ),
        encoding="utf-8",
    )

    exit_code = main([str(request_path), str(results_path)])

    assert exit_code == 1
    assert fake.exported_to is None


def test_cli_rejects_unsupported_operation_per_result(tmp_path, monkeypatch) -> None:
    fake = FakePptxFile()
    monkeypatch.setattr("power_narrator.cli.__main__.PptxFile.open", lambda path: fake)

    request_path = tmp_path / "request.json"
    results_path = tmp_path / "results.json"
    request_path.write_text(
        json.dumps(
            {
                "input": "input.pptx",
                "output": "",
                "ops": [{"op": "export_to", "args": {"output_path": "bad.pptx"}}],
            }
        ),
        encoding="utf-8",
    )

    exit_code = main([str(request_path), str(results_path)])
    results = json.loads(results_path.read_text(encoding="utf-8"))

    assert exit_code == 1
    assert results == {
        "results": [
            {
                "success": False,
                "result": None,
                "message": "Unsupported operation: export_to",
            }
        ]
    }
    assert fake.exported_to is None


def test_cli_reports_request_validation_errors(tmp_path) -> None:
    request_path = tmp_path / "request.json"
    results_path = tmp_path / "results.json"
    request_path.write_text(
        json.dumps({"output": "out.pptx", "ops": []}), encoding="utf-8"
    )

    exit_code = main([str(request_path), str(results_path)])
    results = json.loads(results_path.read_text(encoding="utf-8"))

    assert exit_code == 1
    assert results == {
        "results": [
            {
                "success": False,
                "result": None,
                "message": "Request field 'input' must be a non-empty string",
            }
        ]
    }


def test_cli_rejects_non_strict_argument_types(tmp_path, monkeypatch) -> None:
    fake = FakePptxFile()
    monkeypatch.setattr("power_narrator.cli.__main__.PptxFile.open", lambda path: fake)

    request_path = tmp_path / "request.json"
    results_path = tmp_path / "results.json"
    request_path.write_text(
        json.dumps(
            {
                "input": "input.pptx",
                "output": "output.pptx",
                "ops": [
                    {
                        "op": "set_slide_notes",
                        "args": {"slide_index": "0", "notes": "updated"},
                    },
                    {
                        "op": "save_audio_for_slide",
                        "args": {"slide_index": 0, "mp3_path": 123},
                    },
                ],
            }
        ),
        encoding="utf-8",
    )

    exit_code = main([str(request_path), str(results_path)])
    results = json.loads(results_path.read_text(encoding="utf-8"))

    assert exit_code == 1
    assert results == {
        "results": [
            {
                "success": False,
                "result": None,
                "message": "set_slide_notes() argument 'slide_index' must be int",
            },
            {
                "success": False,
                "result": None,
                "message": "save_audio_for_slide() argument 'mp3_path' must be str",
            },
        ]
    }
    assert fake.exported_to is None


def test_cli_reports_missing_required_argument(tmp_path, monkeypatch) -> None:
    fake = FakePptxFile()
    monkeypatch.setattr("power_narrator.cli.__main__.PptxFile.open", lambda path: fake)

    request_path = tmp_path / "request.json"
    results_path = tmp_path / "results.json"
    request_path.write_text(
        json.dumps(
            {
                "input": "input.pptx",
                "output": "output.pptx",
                "ops": [
                    {"op": "set_slide_notes", "args": {"slide_index": 0}},
                ],
            }
        ),
        encoding="utf-8",
    )

    exit_code = main([str(request_path), str(results_path)])
    results = json.loads(results_path.read_text(encoding="utf-8"))

    assert exit_code == 1
    assert results == {
        "results": [
            {
                "success": False,
                "result": None,
                "message": "set_slide_notes() missing required argument 'notes'",
            }
        ]
    }
    assert fake.exported_to is None


def test_cli_reports_unexpected_argument(tmp_path, monkeypatch) -> None:
    fake = FakePptxFile()
    monkeypatch.setattr("power_narrator.cli.__main__.PptxFile.open", lambda path: fake)

    request_path = tmp_path / "request.json"
    results_path = tmp_path / "results.json"
    request_path.write_text(
        json.dumps(
            {
                "input": "input.pptx",
                "output": "output.pptx",
                "ops": [
                    {
                        "op": "set_slide_notes",
                        "args": {
                            "slide_index": 0,
                            "notes": "updated",
                            "extra": True,
                        },
                    },
                ],
            }
        ),
        encoding="utf-8",
    )

    exit_code = main([str(request_path), str(results_path)])
    results = json.loads(results_path.read_text(encoding="utf-8"))

    assert exit_code == 1
    assert results == {
        "results": [
            {
                "success": False,
                "result": None,
                "message": "set_slide_notes() got unexpected argument 'extra'",
            }
        ]
    }
    assert fake.exported_to is None
