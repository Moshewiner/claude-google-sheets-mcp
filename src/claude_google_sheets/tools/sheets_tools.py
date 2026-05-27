"""Google Sheets API tools for data manipulation and sheet management."""

import json
import logging
import re
from typing import Any, Dict, List, Optional, Tuple

from googleapiclient.errors import HttpError
from mcp.types import TextContent, Tool

from ..auth.oauth_manager import GoogleSheetsAuth
from ..core.exceptions import InvalidRangeError, SheetsAPIError
from ..core.tool_handler import SheetsToolHandler

logger = logging.getLogger(__name__)


def _col_letters_to_index(letters: str) -> int:
    """Convert A1 column letters (e.g. 'A', 'AB') to 0-based index."""
    result = 0
    for ch in letters.upper():
        if not ("A" <= ch <= "Z"):
            raise InvalidRangeError(f"Invalid column letters: {letters}")
        result = result * 26 + (ord(ch) - ord("A") + 1)
    return result - 1


def _parse_a1_range(a1: str) -> Tuple[Optional[str], int, int, int, int]:
    """Parse an A1 range into (sheet_name, start_row, end_row, start_col, end_col).

    Row/col indices are 0-based; end indices are exclusive (GridRange semantics).
    Sheet name may be None if not provided. Supports single cells (A1) and ranges
    (A1:C10). Sheet names may be quoted with single quotes (e.g. 'My Sheet'!A1).
    """
    sheet_name: Optional[str] = None
    cells = a1
    if "!" in a1:
        sheet_part, cells = a1.split("!", 1)
        sheet_name = sheet_part.strip()
        if sheet_name.startswith("'") and sheet_name.endswith("'"):
            sheet_name = sheet_name[1:-1].replace("''", "'")

    match = re.fullmatch(
        r"\s*([A-Za-z]+)(\d+)(?:\s*:\s*([A-Za-z]+)(\d+))?\s*", cells
    )
    if not match:
        raise InvalidRangeError(
            f"Could not parse A1 range: {a1!r}. Expected forms like 'A1' or 'A1:C10'."
        )

    start_col_letters, start_row_str, end_col_letters, end_row_str = match.groups()
    start_col = _col_letters_to_index(start_col_letters)
    start_row = int(start_row_str) - 1
    if end_col_letters is None:
        end_col = start_col + 1
        end_row = start_row + 1
    else:
        end_col = _col_letters_to_index(end_col_letters) + 1
        end_row = int(end_row_str) - 1 + 1

    return sheet_name, start_row, end_row, start_col, end_col


def _resolve_sheet_id(
    sheets_service: Any, spreadsheet_id: str, sheet_name: Optional[str]
) -> int:
    """Look up a tab's numeric sheetId by name. If sheet_name is None, return the first tab's ID."""
    meta = (
        sheets_service.spreadsheets()
        .get(spreadsheetId=spreadsheet_id, fields="sheets.properties(sheetId,title)")
        .execute()
    )
    sheets = meta.get("sheets", [])
    if not sheets:
        raise SheetsAPIError("Spreadsheet has no sheets", 404)
    if sheet_name is None:
        return sheets[0]["properties"]["sheetId"]
    for s in sheets:
        if s["properties"]["title"] == sheet_name:
            return s["properties"]["sheetId"]
    available = ", ".join(s["properties"]["title"] for s in sheets)
    raise InvalidRangeError(
        f"Sheet tab {sheet_name!r} not found. Available tabs: {available}"
    )


class ListSheetTabsHandler(SheetsToolHandler):
    """Handler for listing all tabs (sheets) within a spreadsheet."""

    def __init__(self, auth: GoogleSheetsAuth) -> None:
        super().__init__(
            name="list_sheet_tabs",
            description=(
                "List every tab inside a Google Sheets spreadsheet — returns each tab's "
                "title, numeric sheetId, grid size (rows/columns), hidden flag, and "
                "frozen-row/column counts. Use this to auto-discover tab names when the "
                "caller doesn't know them in advance. Uses only the Sheets API, so it "
                "works even when get_spreadsheet_info (Drive API) fails."
            ),
        )
        self.auth = auth

    def get_tool_definition(self) -> Tool:
        return Tool(
            name=self.name,
            description=self.description,
            inputSchema={
                "type": "object",
                "properties": {
                    "spreadsheet_id": {
                        "type": "string",
                        "description": "The ID of the spreadsheet",
                    },
                    "include_hidden": {
                        "type": "boolean",
                        "description": "Include tabs marked hidden (default: true)",
                        "default": True,
                    },
                },
                "required": ["spreadsheet_id"],
            },
        )

    async def execute(self, arguments: Dict[str, Any]) -> List[TextContent]:
        try:
            self.validate_arguments(arguments, ["spreadsheet_id"])

            spreadsheet_id = arguments["spreadsheet_id"]
            include_hidden = arguments.get("include_hidden", True)

            sheets_service = self.auth.get_sheets_service()

            meta = (
                sheets_service.spreadsheets()
                .get(
                    spreadsheetId=spreadsheet_id,
                    fields=(
                        "spreadsheetId,properties(title,locale,timeZone),"
                        "sheets.properties(sheetId,title,index,sheetType,hidden,"
                        "gridProperties(rowCount,columnCount,frozenRowCount,"
                        "frozenColumnCount),tabColorStyle)"
                    ),
                )
                .execute()
            )

            tabs = []
            for sheet in meta.get("sheets", []):
                props = sheet.get("properties", {})
                if not include_hidden and props.get("hidden"):
                    continue
                tabs.append(
                    {
                        "title": props.get("title"),
                        "sheet_id": props.get("sheetId"),
                        "index": props.get("index"),
                        "sheet_type": props.get("sheetType", "GRID"),
                        "hidden": props.get("hidden", False),
                        "grid": props.get("gridProperties", {}),
                        "tab_color_style": props.get("tabColorStyle"),
                    }
                )

            response_data = {
                "spreadsheet_id": meta.get("spreadsheetId"),
                "spreadsheet_title": meta.get("properties", {}).get("title"),
                "locale": meta.get("properties", {}).get("locale"),
                "time_zone": meta.get("properties", {}).get("timeZone"),
                "tab_count": len(tabs),
                "tabs": tabs,
            }

            return self.format_success_response(
                json.dumps(response_data, indent=2, ensure_ascii=False),
                f"Found {len(tabs)} tab(s) in spreadsheet",
            )

        except HttpError as e:
            if e.resp.status == 404:
                raise SheetsAPIError("Spreadsheet not found", 404)
            else:
                raise SheetsAPIError(f"Sheets API error: {e.reason}", e.resp.status)
        except (InvalidRangeError, SheetsAPIError):
            raise
        except Exception as e:
            return self.format_error_response(e)


class ReadRangeHandler(SheetsToolHandler):
    """Handler for reading data from spreadsheet ranges."""

    def __init__(self, auth: GoogleSheetsAuth) -> None:
        super().__init__(
            name="read_range",
            description="Read data from a specified range in a Google Sheet",
        )
        self.auth = auth

    def get_tool_definition(self) -> Tool:
        return Tool(
            name=self.name,
            description=self.description,
            inputSchema={
                "type": "object",
                "properties": {
                    "spreadsheet_id": {
                        "type": "string",
                        "description": "The ID of the spreadsheet",
                    },
                    "range": {
                        "type": "string",
                        "description": "The A1 notation range to read (e.g., 'Sheet1!A1:C10' or 'A1:C10')",
                    },
                    "value_render_option": {
                        "type": "string",
                        "description": "How values should be represented",
                        "enum": ["FORMATTED_VALUE", "UNFORMATTED_VALUE", "FORMULA"],
                        "default": "FORMATTED_VALUE",
                    },
                    "date_time_render_option": {
                        "type": "string",
                        "description": "How dates should be represented",
                        "enum": ["SERIAL_NUMBER", "FORMATTED_STRING"],
                        "default": "FORMATTED_STRING",
                    },
                },
                "required": ["spreadsheet_id", "range"],
            },
        )

    async def execute(self, arguments: Dict[str, Any]) -> List[TextContent]:
        """Execute the read range operation."""
        try:
            self.validate_arguments(arguments, ["spreadsheet_id", "range"])

            spreadsheet_id = arguments["spreadsheet_id"]
            range_name = arguments["range"]
            value_render_option = arguments.get(
                "value_render_option", "FORMATTED_VALUE"
            )
            date_time_render_option = arguments.get(
                "date_time_render_option", "FORMATTED_STRING"
            )

            sheets_service = self.auth.get_sheets_service()

            result = (
                sheets_service.spreadsheets()
                .values()
                .get(
                    spreadsheetId=spreadsheet_id,
                    range=range_name,
                    valueRenderOption=value_render_option,
                    dateTimeRenderOption=date_time_render_option,
                )
                .execute()
            )

            values = result.get("values", [])

            if not values:
                return self.format_success_response(
                    "No data found in the specified range."
                )

            response_data = {
                "range": result.get("range"),
                "major_dimension": result.get("majorDimension", "ROWS"),
                "row_count": len(values),
                "column_count": max(len(row) for row in values) if values else 0,
                "values": values,
            }

            return self.format_success_response(
                json.dumps(response_data, indent=2),
                f"Read {len(values)} rows from range {range_name}",
            )

        except HttpError as e:
            if e.resp.status == 400:
                raise InvalidRangeError(f"Invalid range: {range_name}")
            elif e.resp.status == 404:
                raise SheetsAPIError("Spreadsheet not found", 404)
            else:
                raise SheetsAPIError(f"Sheets API error: {e.reason}", e.resp.status)
        except Exception as e:
            return self.format_error_response(e)


class WriteRangeHandler(SheetsToolHandler):
    """Handler for writing data to spreadsheet ranges."""

    def __init__(self, auth: GoogleSheetsAuth) -> None:
        super().__init__(
            name="write_range",
            description="Write data to a specified range in a Google Sheet",
        )
        self.auth = auth

    def get_tool_definition(self) -> Tool:
        return Tool(
            name=self.name,
            description=self.description,
            inputSchema={
                "type": "object",
                "properties": {
                    "spreadsheet_id": {
                        "type": "string",
                        "description": "The ID of the spreadsheet",
                    },
                    "range": {
                        "type": "string",
                        "description": "The A1 notation range to write to (e.g., 'Sheet1!A1:C10')",
                    },
                    "values": {
                        "type": "array",
                        "description": "2D array of values to write",
                        "items": {"type": "array", "items": {"type": "string"}},
                    },
                    "value_input_option": {
                        "type": "string",
                        "description": "How input data should be interpreted",
                        "enum": ["RAW", "USER_ENTERED"],
                        "default": "USER_ENTERED",
                    },
                },
                "required": ["spreadsheet_id", "range", "values"],
            },
        )

    async def execute(self, arguments: Dict[str, Any]) -> List[TextContent]:
        """Execute the write range operation."""
        try:
            self.validate_arguments(arguments, ["spreadsheet_id", "range", "values"])

            spreadsheet_id = arguments["spreadsheet_id"]
            range_name = arguments["range"]
            values = arguments["values"]
            value_input_option = arguments.get("value_input_option", "USER_ENTERED")

            sheets_service = self.auth.get_sheets_service()

            body = {"values": values}

            result = (
                sheets_service.spreadsheets()
                .values()
                .update(
                    spreadsheetId=spreadsheet_id,
                    range=range_name,
                    valueInputOption=value_input_option,
                    body=body,
                )
                .execute()
            )

            response_data = {
                "updated_range": result.get("updatedRange"),
                "updated_rows": result.get("updatedRows"),
                "updated_columns": result.get("updatedColumns"),
                "updated_cells": result.get("updatedCells"),
            }

            return self.format_success_response(
                json.dumps(response_data, indent=2),
                f"Successfully updated {result.get('updatedCells', 0)} cells in range {range_name}",
            )

        except HttpError as e:
            if e.resp.status == 400:
                raise InvalidRangeError(f"Invalid range or data: {range_name}")
            elif e.resp.status == 404:
                raise SheetsAPIError("Spreadsheet not found", 404)
            else:
                raise SheetsAPIError(f"Sheets API error: {e.reason}", e.resp.status)
        except Exception as e:
            return self.format_error_response(e)


class AppendDataHandler(SheetsToolHandler):
    """Handler for appending data to a spreadsheet."""

    def __init__(self, auth: GoogleSheetsAuth) -> None:
        super().__init__(
            name="append_data",
            description="Append rows of data to the end of a Google Sheet",
        )
        self.auth = auth

    def get_tool_definition(self) -> Tool:
        return Tool(
            name=self.name,
            description=self.description,
            inputSchema={
                "type": "object",
                "properties": {
                    "spreadsheet_id": {
                        "type": "string",
                        "description": "The ID of the spreadsheet",
                    },
                    "range": {
                        "type": "string",
                        "description": "The A1 notation range indicating the sheet and columns (e.g., 'Sheet1!A:C')",
                    },
                    "values": {
                        "type": "array",
                        "description": "2D array of values to append",
                        "items": {"type": "array", "items": {"type": "string"}},
                    },
                    "value_input_option": {
                        "type": "string",
                        "description": "How input data should be interpreted",
                        "enum": ["RAW", "USER_ENTERED"],
                        "default": "USER_ENTERED",
                    },
                    "insert_data_option": {
                        "type": "string",
                        "description": "How data should be inserted",
                        "enum": ["OVERWRITE", "INSERT_ROWS"],
                        "default": "INSERT_ROWS",
                    },
                },
                "required": ["spreadsheet_id", "range", "values"],
            },
        )

    async def execute(self, arguments: Dict[str, Any]) -> List[TextContent]:
        """Execute the append data operation."""
        try:
            self.validate_arguments(arguments, ["spreadsheet_id", "range", "values"])

            spreadsheet_id = arguments["spreadsheet_id"]
            range_name = arguments["range"]
            values = arguments["values"]
            value_input_option = arguments.get("value_input_option", "USER_ENTERED")
            insert_data_option = arguments.get("insert_data_option", "INSERT_ROWS")

            sheets_service = self.auth.get_sheets_service()

            body = {"values": values}

            result = (
                sheets_service.spreadsheets()
                .values()
                .append(
                    spreadsheetId=spreadsheet_id,
                    range=range_name,
                    valueInputOption=value_input_option,
                    insertDataOption=insert_data_option,
                    body=body,
                )
                .execute()
            )

            response_data = {
                "spreadsheet_id": result.get("spreadsheetId"),
                "table_range": result.get("tableRange"),
                "updates": result.get("updates", {}),
            }

            return self.format_success_response(
                json.dumps(response_data, indent=2),
                f"Successfully appended {len(values)} rows to {range_name}",
            )

        except HttpError as e:
            if e.resp.status == 400:
                raise InvalidRangeError(f"Invalid range or data: {range_name}")
            elif e.resp.status == 404:
                raise SheetsAPIError("Spreadsheet not found", 404)
            else:
                raise SheetsAPIError(f"Sheets API error: {e.reason}", e.resp.status)
        except Exception as e:
            return self.format_error_response(e)


class ClearRangeHandler(SheetsToolHandler):
    """Handler for clearing data from spreadsheet ranges."""

    def __init__(self, auth: GoogleSheetsAuth) -> None:
        super().__init__(
            name="clear_range",
            description="Clear data from a specified range in a Google Sheet",
        )
        self.auth = auth

    def get_tool_definition(self) -> Tool:
        return Tool(
            name=self.name,
            description=self.description,
            inputSchema={
                "type": "object",
                "properties": {
                    "spreadsheet_id": {
                        "type": "string",
                        "description": "The ID of the spreadsheet",
                    },
                    "range": {
                        "type": "string",
                        "description": "The A1 notation range to clear (e.g., 'Sheet1!A1:C10')",
                    },
                },
                "required": ["spreadsheet_id", "range"],
            },
        )

    async def execute(self, arguments: Dict[str, Any]) -> List[TextContent]:
        """Execute the clear range operation."""
        try:
            self.validate_arguments(arguments, ["spreadsheet_id", "range"])

            spreadsheet_id = arguments["spreadsheet_id"]
            range_name = arguments["range"]

            sheets_service = self.auth.get_sheets_service()

            result = (
                sheets_service.spreadsheets()
                .values()
                .clear(spreadsheetId=spreadsheet_id, range=range_name, body={})
                .execute()
            )

            response_data = {
                "cleared_range": result.get("clearedRange"),
                "spreadsheet_id": result.get("spreadsheetId"),
            }

            return self.format_success_response(
                json.dumps(response_data, indent=2),
                f"Successfully cleared range {range_name}",
            )

        except HttpError as e:
            if e.resp.status == 400:
                raise InvalidRangeError(f"Invalid range: {range_name}")
            elif e.resp.status == 404:
                raise SheetsAPIError("Spreadsheet not found", 404)
            else:
                raise SheetsAPIError(f"Sheets API error: {e.reason}", e.resp.status)
        except Exception as e:
            return self.format_error_response(e)


class AddCellNoteHandler(SheetsToolHandler):
    """Handler for attaching a hover note (yellow-triangle audit note) to a cell or range."""

    def __init__(self, auth: GoogleSheetsAuth) -> None:
        super().__init__(
            name="add_cell_note",
            description=(
                "Attach a hover note (the yellow-triangle annotation visible on cell hover) "
                "to a single cell or a rectangular range. Use this to audit a calculated "
                "value by explaining where it came from. If the range covers multiple cells, "
                "the same note is applied to every cell in the range. Pass an empty string "
                "to clear an existing note."
            ),
        )
        self.auth = auth

    def get_tool_definition(self) -> Tool:
        return Tool(
            name=self.name,
            description=self.description,
            inputSchema={
                "type": "object",
                "properties": {
                    "spreadsheet_id": {
                        "type": "string",
                        "description": "The ID of the spreadsheet",
                    },
                    "range": {
                        "type": "string",
                        "description": (
                            "A1 notation for the target cell or range "
                            "(e.g. 'Sheet1!I8' or \"'תזרים'!I8:I50\")."
                        ),
                    },
                    "note": {
                        "type": "string",
                        "description": (
                            "The note text to attach. Pass an empty string to clear "
                            "any existing note on the target cells."
                        ),
                    },
                },
                "required": ["spreadsheet_id", "range", "note"],
            },
        )

    async def execute(self, arguments: Dict[str, Any]) -> List[TextContent]:
        try:
            self.validate_arguments(arguments, ["spreadsheet_id", "range", "note"])

            spreadsheet_id = arguments["spreadsheet_id"]
            range_name = arguments["range"]
            note_text = arguments["note"]

            sheet_name, start_row, end_row, start_col, end_col = _parse_a1_range(
                range_name
            )

            sheets_service = self.auth.get_sheets_service()
            sheet_id = _resolve_sheet_id(sheets_service, spreadsheet_id, sheet_name)

            request_body = {
                "requests": [
                    {
                        "repeatCell": {
                            "range": {
                                "sheetId": sheet_id,
                                "startRowIndex": start_row,
                                "endRowIndex": end_row,
                                "startColumnIndex": start_col,
                                "endColumnIndex": end_col,
                            },
                            "cell": {"note": note_text},
                            "fields": "note",
                        }
                    }
                ]
            }

            result = (
                sheets_service.spreadsheets()
                .batchUpdate(spreadsheetId=spreadsheet_id, body=request_body)
                .execute()
            )

            cell_count = (end_row - start_row) * (end_col - start_col)
            response_data = {
                "spreadsheet_id": result.get("spreadsheetId"),
                "range": range_name,
                "sheet_id": sheet_id,
                "cells_updated": cell_count,
                "note_cleared": note_text == "",
            }

            return self.format_success_response(
                json.dumps(response_data, indent=2, ensure_ascii=False),
                f"Applied note to {cell_count} cell(s) in {range_name}",
            )

        except HttpError as e:
            if e.resp.status == 400:
                raise InvalidRangeError(f"Invalid range or request: {range_name}")
            elif e.resp.status == 404:
                raise SheetsAPIError("Spreadsheet not found", 404)
            else:
                raise SheetsAPIError(f"Sheets API error: {e.reason}", e.resp.status)
        except (InvalidRangeError, SheetsAPIError):
            raise
        except Exception as e:
            return self.format_error_response(e)


def _to_user_entered_value(value: Any) -> Dict[str, Any]:
    """Translate a Python value into a Sheets API `userEnteredValue` payload."""
    if value is None:
        return {}
    if isinstance(value, bool):
        return {"boolValue": value}
    if isinstance(value, (int, float)):
        return {"numberValue": float(value)}
    if isinstance(value, str):
        if value.startswith("="):
            return {"formulaValue": value}
        return {"stringValue": value}
    return {"stringValue": str(value)}


def _parse_color(color: Any) -> Optional[Dict[str, float]]:
    """Accept either '#rrggbb' / '#rgb' hex or {red,green,blue,alpha?} dict (0-1 floats)."""
    if color is None:
        return None
    if isinstance(color, dict):
        out = {}
        for k in ("red", "green", "blue", "alpha"):
            if k in color:
                v = float(color[k])
                if v > 1.0:
                    v = v / 255.0
                out[k] = max(0.0, min(1.0, v))
        return out
    if isinstance(color, str):
        s = color.strip().lstrip("#")
        if len(s) == 3:
            s = "".join(ch * 2 for ch in s)
        if len(s) == 6:
            try:
                r = int(s[0:2], 16) / 255.0
                g = int(s[2:4], 16) / 255.0
                b = int(s[4:6], 16) / 255.0
                return {"red": r, "green": g, "blue": b}
            except ValueError:
                pass
    raise InvalidRangeError(f"Could not parse color: {color!r}")


class BatchUpdateCellsHandler(SheetsToolHandler):
    """Write many cells (optionally with hover notes) in a single API call."""

    def __init__(self, auth: GoogleSheetsAuth) -> None:
        super().__init__(
            name="batch_update_cells",
            description=(
                "Update many individual cells in one Sheets API call. Each update sets a "
                "value and optionally a hover note (the yellow-triangle audit note). Use "
                "this instead of looping write_range + add_cell_note when updating many "
                "cells — one batched request is dramatically faster and avoids rate "
                "limits. Values can be numbers, booleans, strings, or formulas (strings "
                "beginning with '=')."
            ),
        )
        self.auth = auth

    def get_tool_definition(self) -> Tool:
        return Tool(
            name=self.name,
            description=self.description,
            inputSchema={
                "type": "object",
                "properties": {
                    "spreadsheet_id": {
                        "type": "string",
                        "description": "The ID of the spreadsheet",
                    },
                    "updates": {
                        "type": "array",
                        "description": (
                            "List of cell updates. Each item must include `range` (A1 "
                            "notation pointing at a single cell, e.g. \"'תזרים'!J8\") "
                            "and `value`. Optionally `note` (string; pass empty string "
                            "to clear an existing note)."
                        ),
                        "items": {
                            "type": "object",
                            "properties": {
                                "range": {
                                    "type": "string",
                                    "description": "A1 notation for a single cell",
                                },
                                "value": {
                                    "description": (
                                        "Number, boolean, string, or formula (string "
                                        "starting with '='). Null clears the cell value."
                                    )
                                },
                                "note": {
                                    "type": "string",
                                    "description": (
                                        "Optional hover note. Pass empty string to clear."
                                    ),
                                },
                            },
                            "required": ["range", "value"],
                        },
                    },
                },
                "required": ["spreadsheet_id", "updates"],
            },
        )

    async def execute(self, arguments: Dict[str, Any]) -> List[TextContent]:
        try:
            self.validate_arguments(arguments, ["spreadsheet_id", "updates"])

            spreadsheet_id = arguments["spreadsheet_id"]
            updates = arguments["updates"]
            if not isinstance(updates, list) or not updates:
                raise InvalidRangeError("`updates` must be a non-empty list")

            sheets_service = self.auth.get_sheets_service()

            # Resolve sheet IDs once per distinct tab name.
            sheet_id_cache: Dict[Optional[str], int] = {}
            requests = []

            for idx, entry in enumerate(updates):
                if not isinstance(entry, dict) or "range" not in entry or "value" not in entry:
                    raise InvalidRangeError(
                        f"updates[{idx}] must be an object with `range` and `value`"
                    )

                a1 = entry["range"]
                sheet_name, start_row, end_row, start_col, end_col = _parse_a1_range(a1)
                if (end_row - start_row) != 1 or (end_col - start_col) != 1:
                    raise InvalidRangeError(
                        f"updates[{idx}] range {a1!r} must point to a single cell"
                    )

                if sheet_name not in sheet_id_cache:
                    sheet_id_cache[sheet_name] = _resolve_sheet_id(
                        sheets_service, spreadsheet_id, sheet_name
                    )
                sheet_id = sheet_id_cache[sheet_name]

                cell: Dict[str, Any] = {
                    "userEnteredValue": _to_user_entered_value(entry["value"])
                }
                fields_parts = ["userEnteredValue"]
                if "note" in entry:
                    cell["note"] = entry["note"]
                    fields_parts.append("note")

                requests.append(
                    {
                        "updateCells": {
                            "range": {
                                "sheetId": sheet_id,
                                "startRowIndex": start_row,
                                "endRowIndex": end_row,
                                "startColumnIndex": start_col,
                                "endColumnIndex": end_col,
                            },
                            "rows": [{"values": [cell]}],
                            "fields": ",".join(fields_parts),
                        }
                    }
                )

            result = (
                sheets_service.spreadsheets()
                .batchUpdate(
                    spreadsheetId=spreadsheet_id, body={"requests": requests}
                )
                .execute()
            )

            response_data = {
                "spreadsheet_id": result.get("spreadsheetId"),
                "updates_applied": len(requests),
                "reply_count": len(result.get("replies", [])),
            }

            return self.format_success_response(
                json.dumps(response_data, indent=2, ensure_ascii=False),
                f"Applied {len(requests)} cell update(s) in one batch",
            )

        except HttpError as e:
            if e.resp.status == 400:
                raise InvalidRangeError(f"Invalid batch update request: {e.reason}")
            elif e.resp.status == 404:
                raise SheetsAPIError("Spreadsheet not found", 404)
            else:
                raise SheetsAPIError(f"Sheets API error: {e.reason}", e.resp.status)
        except (InvalidRangeError, SheetsAPIError):
            raise
        except Exception as e:
            return self.format_error_response(e)


class BatchReadRangesHandler(SheetsToolHandler):
    """Read multiple ranges in a single Sheets API call."""

    def __init__(self, auth: GoogleSheetsAuth) -> None:
        super().__init__(
            name="batch_read_ranges",
            description=(
                "Read multiple A1 ranges from one spreadsheet in a single API call. "
                "Use this when you need data from several tabs or several disjoint "
                "ranges — much faster than calling read_range repeatedly."
            ),
        )
        self.auth = auth

    def get_tool_definition(self) -> Tool:
        return Tool(
            name=self.name,
            description=self.description,
            inputSchema={
                "type": "object",
                "properties": {
                    "spreadsheet_id": {
                        "type": "string",
                        "description": "The ID of the spreadsheet",
                    },
                    "ranges": {
                        "type": "array",
                        "description": "List of A1 ranges to read",
                        "items": {"type": "string"},
                    },
                    "value_render_option": {
                        "type": "string",
                        "description": "How values should be represented",
                        "enum": ["FORMATTED_VALUE", "UNFORMATTED_VALUE", "FORMULA"],
                        "default": "FORMATTED_VALUE",
                    },
                    "date_time_render_option": {
                        "type": "string",
                        "description": "How dates should be represented",
                        "enum": ["SERIAL_NUMBER", "FORMATTED_STRING"],
                        "default": "FORMATTED_STRING",
                    },
                },
                "required": ["spreadsheet_id", "ranges"],
            },
        )

    async def execute(self, arguments: Dict[str, Any]) -> List[TextContent]:
        try:
            self.validate_arguments(arguments, ["spreadsheet_id", "ranges"])

            spreadsheet_id = arguments["spreadsheet_id"]
            ranges = arguments["ranges"]
            if not isinstance(ranges, list) or not ranges:
                raise InvalidRangeError("`ranges` must be a non-empty list")

            value_render_option = arguments.get(
                "value_render_option", "FORMATTED_VALUE"
            )
            date_time_render_option = arguments.get(
                "date_time_render_option", "FORMATTED_STRING"
            )

            sheets_service = self.auth.get_sheets_service()

            result = (
                sheets_service.spreadsheets()
                .values()
                .batchGet(
                    spreadsheetId=spreadsheet_id,
                    ranges=ranges,
                    valueRenderOption=value_render_option,
                    dateTimeRenderOption=date_time_render_option,
                )
                .execute()
            )

            ranges_out = []
            for vr in result.get("valueRanges", []):
                values = vr.get("values", [])
                ranges_out.append(
                    {
                        "range": vr.get("range"),
                        "major_dimension": vr.get("majorDimension", "ROWS"),
                        "row_count": len(values),
                        "column_count": max((len(r) for r in values), default=0),
                        "values": values,
                    }
                )

            response_data = {
                "spreadsheet_id": result.get("spreadsheetId"),
                "range_count": len(ranges_out),
                "ranges": ranges_out,
            }

            return self.format_success_response(
                json.dumps(response_data, indent=2, ensure_ascii=False),
                f"Read {len(ranges_out)} range(s) in one batch",
            )

        except HttpError as e:
            if e.resp.status == 400:
                raise InvalidRangeError(f"Invalid batch read request: {e.reason}")
            elif e.resp.status == 404:
                raise SheetsAPIError("Spreadsheet not found", 404)
            else:
                raise SheetsAPIError(f"Sheets API error: {e.reason}", e.resp.status)
        except (InvalidRangeError, SheetsAPIError):
            raise
        except Exception as e:
            return self.format_error_response(e)


class SetCellFormatHandler(SheetsToolHandler):
    """Apply visual formatting (color, bold, number format) to a cell range."""

    def __init__(self, auth: GoogleSheetsAuth) -> None:
        super().__init__(
            name="set_cell_format",
            description=(
                "Apply visual formatting to a cell or rectangular range — background "
                "color, text color, bold/italic, and/or number format pattern. Useful "
                "for highlighting anomalies in an audit (e.g., red background when an "
                "actual exceeds the planned amount). Only the fields you provide are "
                "modified; other formatting is preserved."
            ),
        )
        self.auth = auth

    def get_tool_definition(self) -> Tool:
        return Tool(
            name=self.name,
            description=self.description,
            inputSchema={
                "type": "object",
                "properties": {
                    "spreadsheet_id": {
                        "type": "string",
                        "description": "The ID of the spreadsheet",
                    },
                    "range": {
                        "type": "string",
                        "description": "A1 notation for a cell or rectangular range",
                    },
                    "background_color": {
                        "description": (
                            "Cell background. Hex like '#ffcccc' or '{red,green,blue}' "
                            "with floats in 0-1 (or ints 0-255). Omit to leave unchanged."
                        )
                    },
                    "text_color": {
                        "description": "Text color, same format as background_color."
                    },
                    "bold": {
                        "type": "boolean",
                        "description": "Set bold on/off. Omit to leave unchanged.",
                    },
                    "italic": {
                        "type": "boolean",
                        "description": "Set italic on/off. Omit to leave unchanged.",
                    },
                    "number_format": {
                        "type": "string",
                        "description": (
                            "Number format pattern (e.g., '#,##0.00', '\"₪\"#,##0.00', "
                            "'0.00%'). Omit to leave unchanged."
                        ),
                    },
                    "number_format_type": {
                        "type": "string",
                        "description": "Format type when setting number_format.",
                        "enum": [
                            "NUMBER",
                            "PERCENT",
                            "CURRENCY",
                            "DATE",
                            "TIME",
                            "DATE_TIME",
                            "SCIENTIFIC",
                            "TEXT",
                        ],
                        "default": "NUMBER",
                    },
                },
                "required": ["spreadsheet_id", "range"],
            },
        )

    async def execute(self, arguments: Dict[str, Any]) -> List[TextContent]:
        try:
            self.validate_arguments(arguments, ["spreadsheet_id", "range"])

            spreadsheet_id = arguments["spreadsheet_id"]
            range_name = arguments["range"]

            sheet_name, start_row, end_row, start_col, end_col = _parse_a1_range(
                range_name
            )

            user_format: Dict[str, Any] = {}
            field_parts: List[str] = []

            bg = _parse_color(arguments.get("background_color"))
            if bg is not None:
                user_format["backgroundColorStyle"] = {"rgbColor": bg}
                field_parts.append("userEnteredFormat.backgroundColorStyle")

            txt = _parse_color(arguments.get("text_color"))
            text_format: Dict[str, Any] = {}
            text_field_parts: List[str] = []
            if txt is not None:
                text_format["foregroundColorStyle"] = {"rgbColor": txt}
                text_field_parts.append("foregroundColorStyle")
            if "bold" in arguments:
                text_format["bold"] = bool(arguments["bold"])
                text_field_parts.append("bold")
            if "italic" in arguments:
                text_format["italic"] = bool(arguments["italic"])
                text_field_parts.append("italic")
            if text_format:
                user_format["textFormat"] = text_format
                field_parts.extend(
                    f"userEnteredFormat.textFormat.{f}" for f in text_field_parts
                )

            if "number_format" in arguments and arguments["number_format"]:
                user_format["numberFormat"] = {
                    "type": arguments.get("number_format_type", "NUMBER"),
                    "pattern": arguments["number_format"],
                }
                field_parts.append("userEnteredFormat.numberFormat")

            if not user_format:
                raise InvalidRangeError(
                    "set_cell_format requires at least one of: background_color, "
                    "text_color, bold, italic, number_format"
                )

            sheets_service = self.auth.get_sheets_service()
            sheet_id = _resolve_sheet_id(sheets_service, spreadsheet_id, sheet_name)

            request_body = {
                "requests": [
                    {
                        "repeatCell": {
                            "range": {
                                "sheetId": sheet_id,
                                "startRowIndex": start_row,
                                "endRowIndex": end_row,
                                "startColumnIndex": start_col,
                                "endColumnIndex": end_col,
                            },
                            "cell": {"userEnteredFormat": user_format},
                            "fields": ",".join(field_parts),
                        }
                    }
                ]
            }

            result = (
                sheets_service.spreadsheets()
                .batchUpdate(spreadsheetId=spreadsheet_id, body=request_body)
                .execute()
            )

            cell_count = (end_row - start_row) * (end_col - start_col)
            response_data = {
                "spreadsheet_id": result.get("spreadsheetId"),
                "range": range_name,
                "sheet_id": sheet_id,
                "cells_updated": cell_count,
                "fields_changed": field_parts,
            }

            return self.format_success_response(
                json.dumps(response_data, indent=2, ensure_ascii=False),
                f"Formatted {cell_count} cell(s) in {range_name}",
            )

        except HttpError as e:
            if e.resp.status == 400:
                raise InvalidRangeError(f"Invalid format request: {e.reason}")
            elif e.resp.status == 404:
                raise SheetsAPIError("Spreadsheet not found", 404)
            else:
                raise SheetsAPIError(f"Sheets API error: {e.reason}", e.resp.status)
        except (InvalidRangeError, SheetsAPIError):
            raise
        except Exception as e:
            return self.format_error_response(e)


# Registry of all sheets tool handlers
SHEETS_HANDLERS = [
    ListSheetTabsHandler,
    ReadRangeHandler,
    BatchReadRangesHandler,
    WriteRangeHandler,
    BatchUpdateCellsHandler,
    AppendDataHandler,
    ClearRangeHandler,
    AddCellNoteHandler,
    SetCellFormatHandler,
]
