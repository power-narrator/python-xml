"""CLI entrypoint for JSON-driven PPTX operations."""

import argparse
from dataclasses import dataclass
import inspect
import json
import sys
from pathlib import Path
from typing import Any, TypedDict

from .pptx_file import PptxFile

ALLOWED_OPERATIONS = {
    name
    for name, member in inspect.getmembers(PptxFile, predicate=callable)
    if not name.startswith("_") and name not in {"close", "export_to", "open"}
}


class OperationPayload(TypedDict, total=False):
    """Raw JSON operation payload."""

    op: str
    args: list[Any]


class RequestPayload(TypedDict, total=False):
    """Raw JSON request payload."""

    input: str
    output: str
    ops: list[OperationPayload]


class OperationResultPayload(TypedDict):
    """JSON-serializable per-operation result payload."""

    success: bool
    result: Any
    message: str


class ResultsPayload(TypedDict):
    """JSON-serializable batch results payload."""

    results: list[OperationResultPayload]


@dataclass(frozen=True)
class Operation:
    """Validated operation request."""

    op: str
    args: list[Any]


@dataclass(frozen=True)
class Request:
    """Validated batch request."""

    input_path: Path
    output_path: Path | None
    operations: list[Operation]


def _jsonable(value: Any) -> Any:
    """Convert a value into a JSON-serializable structure.

    Args:
        value: Value returned from an operation.

    Returns:
        JSON-serializable value.
    """
    if isinstance(value, Path):
        return str(value)

    if isinstance(value, tuple):
        return [_jsonable(item) for item in value]

    if isinstance(value, list):
        return [_jsonable(item) for item in value]

    if isinstance(value, dict):
        return {str(key): _jsonable(item) for key, item in value.items()}

    return value


def _coerce_arg(method_name: str, parameter: inspect.Parameter, arg: Any) -> Any:
    """Coerce a single JSON argument to a strict annotated type.

    Args:
        method_name: Name of the target ``PptxFile`` method.
        parameter: Target parameter definition.
        arg: Raw JSON argument value.

    Returns:
        Strictly validated argument value.

    Raises:
        TypeError: If the argument does not match the expected annotation or the
            annotation is unsupported by the CLI.
    """
    annotation = parameter.annotation
    expected_name = getattr(annotation, "__name__", repr(annotation))

    if annotation is inspect._empty:
        return arg

    if annotation is Path:
        if type(arg) is not str:
            raise TypeError(f"{method_name}() argument '{parameter.name}' must be str")

        return Path(arg)

    if annotation is int:
        if type(arg) is not int:
            raise TypeError(f"{method_name}() argument '{parameter.name}' must be int")

        return arg

    if annotation is str:
        if type(arg) is not str:
            raise TypeError(f"{method_name}() argument '{parameter.name}' must be str")

        return arg

    if annotation is bool:
        if type(arg) is not bool:
            raise TypeError(f"{method_name}() argument '{parameter.name}' must be bool")

        return arg

    if annotation is float:
        if type(arg) is not float:
            raise TypeError(
                f"{method_name}() argument '{parameter.name}' must be float"
            )

        return arg

    raise TypeError(
        f"{method_name}() argument '{parameter.name}' uses unsupported annotation {expected_name}"
    )


def _coerce_args(method_name: str, args: list[Any]) -> list[Any]:
    """Coerce JSON operation arguments to method parameter types.

    Args:
        method_name: Name of the target ``PptxFile`` method.
        args: Positional arguments from the request payload.

    Returns:
        List of strictly validated arguments.

    Raises:
        TypeError: If the argument count does not match the method signature
            or any argument fails strict type validation.
    """
    method = getattr(PptxFile, method_name)
    parameters = list(inspect.signature(method).parameters.values())[1:]

    if len(args) != len(parameters):
        raise TypeError(
            f"{method_name}() takes {len(parameters)} positional argument(s) but {len(args)} were given"
        )

    coerced: list[Any] = []

    for arg, parameter in zip(args, parameters, strict=True):
        coerced.append(_coerce_arg(method_name, parameter, arg))

    return coerced


def _read_request(path: Path) -> RequestPayload:
    """Load and validate the top-level JSON payload type.

    Args:
        path: Path to the request JSON file.

    Returns:
        Request payload as a dictionary.

    Raises:
        ValueError: If the parsed JSON root is not an object.
    """
    with path.open(encoding="utf-8") as handle:
        payload = json.load(handle)

    if not isinstance(payload, dict):
        raise ValueError("Request JSON must be an object")

    return RequestPayload(**payload)


def _validate_request(payload: RequestPayload) -> Request:
    """Validate and normalize CLI request payload fields.

    Args:
        payload: Raw request payload loaded from JSON.

    Returns:
        Validated request object.

    Raises:
        ValueError: If any required field is missing or has an invalid type.
    """
    input_value = payload.get("input")
    output_value = payload.get("output")
    ops = payload.get("ops")

    if not isinstance(input_value, str) or not input_value.strip():
        raise ValueError("Request field 'input' must be a non-empty string")

    if output_value is not None and not isinstance(output_value, str):
        raise ValueError("Request field 'output' must be a string when provided")

    if not isinstance(ops, list):
        raise ValueError("Request field 'ops' must be an array")

    normalized_ops: list[Operation] = []

    for index, op in enumerate(ops):
        if not isinstance(op, dict):
            raise ValueError(f"Operation at index {index} must be an object")

        op_name = op.get("op")
        args = op.get("args", [])

        if not isinstance(op_name, str) or not op_name:
            raise ValueError(
                f"Operation at index {index} must have a non-empty 'op' string"
            )

        if not isinstance(args, list):
            raise ValueError(f"Operation '{op_name}' args must be an array")

        normalized_ops.append(Operation(op=op_name, args=args))

    output_path = output_value.strip() if isinstance(output_value, str) else None
    return Request(
        input_path=Path(input_value),
        output_path=Path(output_path) if output_path else None,
        operations=normalized_ops,
    )


def _execute_request(payload: RequestPayload) -> ResultsPayload:
    """Execute a batch of PPTX operations from a request payload.

    Args:
        payload: Request object containing ``input``, ``output``, and ``ops``.

    Returns:
        Result object with one per-operation result entry.

    Raises:
        ValueError: If the request payload is invalid.
        FileNotFoundError: If the input PPTX file does not exist.
        InvalidPptxError: If the input file is not a valid PPTX.
        RelsNotFoundError: If presentation relationships are missing.
    """
    request = _validate_request(payload)
    results: list[OperationResultPayload] = []
    should_export = True

    with PptxFile.open(request.input_path) as pptx_file:
        for operation in request.operations:
            op_name = operation.op

            if op_name not in ALLOWED_OPERATIONS:
                should_export = False
                results.append(
                    {
                        "success": False,
                        "result": None,
                        "message": f"Unsupported operation: {op_name}",
                    }
                )
                continue

            try:
                method = getattr(pptx_file, op_name)
                args = _coerce_args(op_name, operation.args)
                result = method(*args)
            except Exception as exc:
                should_export = False
                results.append({"success": False, "result": None, "message": str(exc)})
                continue

            results.append(
                {"success": True, "result": _jsonable(result), "message": ""}
            )

        if should_export and request.output_path is not None:
            pptx_file.export_to(request.output_path)

    return ResultsPayload(results=results)


def build_parser() -> argparse.ArgumentParser:
    """Create the command-line argument parser.

    Returns:
        Configured parser for request and results JSON paths.
    """
    parser = argparse.ArgumentParser(
        description="Execute JSON-defined PPTX operations against a PowerPoint file."
    )
    parser.add_argument("request_json", type=Path, help="Path to request JSON file")
    parser.add_argument(
        "results_json",
        type=Path,
        help="Path to write operation results JSON",
    )
    return parser


def main(argv: list[str] | None = None) -> int:
    """Run the CLI command.

    Args:
        argv: Optional argument list; when ``None``, values are read from
            ``sys.argv``.

    Returns:
        Process exit code (``0`` when all operations succeed, otherwise ``1``).
    """
    args = build_parser().parse_args(argv)

    try:
        payload = _read_request(args.request_json)
        results = _execute_request(payload)
    except Exception as exc:
        failure = {"results": [{"success": False, "result": None, "message": str(exc)}]}
        output = json.dumps(failure, indent=2)

        if args.results_json is not None:
            args.results_json.write_text(output + "\n", encoding="utf-8")

        print(output)
        return 1

    output = json.dumps(results, indent=2)

    if args.results_json is not None:
        args.results_json.write_text(output + "\n", encoding="utf-8")

    print(output)

    return 0 if all(result["success"] for result in results["results"]) else 1


if __name__ == "__main__":
    sys.exit(main())
