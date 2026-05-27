"""Threaded-comment tools for Google Sheets (Drive API comments + replies).

Unlike `add_cell_note` (the yellow-triangle hover note), these are the
discussion-style comments shown in the Comments pane, with threaded replies
and resolution. They live on the underlying Drive file, not the Sheet grid,
but can be anchored to a specific cell or range so the comment shows up on
that cell in the Sheets UI.
"""

import json
import logging
from typing import Any, Dict, List, Optional

from googleapiclient.errors import HttpError
from mcp.types import TextContent, Tool

from ..auth.oauth_manager import GoogleSheetsAuth
from ..core.exceptions import DriveAPIError, InvalidRangeError, SheetsAPIError
from ..core.tool_handler import SheetsToolHandler
from .sheets_tools import _parse_a1_range, _resolve_sheet_id

logger = logging.getLogger(__name__)

# Fields returned for comment objects. Drive v3 requires explicit fields.
_COMMENT_FIELDS = (
    "id,content,htmlContent,createdTime,modifiedTime,resolved,deleted,"
    "anchor,quotedFileContent,author(displayName,emailAddress,me),"
    "replies(id,content,createdTime,modifiedTime,action,deleted,"
    "author(displayName,emailAddress,me))"
)
_COMMENT_LIST_FIELDS = f"nextPageToken,comments({_COMMENT_FIELDS})"
_REPLY_FIELDS = (
    "id,content,createdTime,modifiedTime,action,deleted,"
    "author(displayName,emailAddress,me)"
)


def _build_sheets_anchor(
    sheets_service: Any,
    spreadsheet_id: str,
    range_a1: str,
) -> str:
    """Build the Drive-comment anchor JSON for a Sheets cell or range.

    Google's anchor format for Sheets comments (undocumented but stable):
        {"type":"workbook-range","uid":1,"range":{
            "rangeType":"GRID_RANGE","sheetId":<id>,
            "startRow":<r>,"endRow":<r+1>,
            "startColumn":<c>,"endColumn":<c+1>}}
    """
    sheet_name, start_row, end_row, start_col, end_col = _parse_a1_range(range_a1)
    sheet_id = _resolve_sheet_id(sheets_service, spreadsheet_id, sheet_name)
    anchor = {
        "type": "workbook-range",
        "uid": 1,
        "range": {
            "rangeType": "GRID_RANGE",
            "sheetId": sheet_id,
            "startRow": start_row,
            "endRow": end_row,
            "startColumn": start_col,
            "endColumn": end_col,
        },
    }
    return json.dumps(anchor, separators=(",", ":"))


class AddCommentHandler(SheetsToolHandler):
    """Create a threaded comment on a Google Sheet, optionally anchored to a cell/range."""

    def __init__(self, auth: GoogleSheetsAuth) -> None:
        super().__init__(
            name="add_comment",
            description=(
                "Create a threaded comment (discussion-style, shown in the Comments pane) "
                "on a Google Sheet. If `range` is provided, the comment is anchored to that "
                "cell or rectangular range and appears on that cell in the UI; otherwise it "
                "is attached to the file as an unanchored comment. Returns the new comment's "
                "ID so it can later be replied to or resolved."
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
                    "comment": {
                        "type": "string",
                        "description": "The comment text",
                    },
                    "range": {
                        "type": "string",
                        "description": (
                            "Optional A1 notation for the cell or range to anchor the comment to "
                            "(e.g. 'Sheet1!B4' or \"'תזרים'!I8:I12\"). Omit for an unanchored "
                            "file-level comment."
                        ),
                    },
                },
                "required": ["spreadsheet_id", "comment"],
            },
        )

    async def execute(self, arguments: Dict[str, Any]) -> List[TextContent]:
        try:
            self.validate_arguments(arguments, ["spreadsheet_id", "comment"])

            spreadsheet_id = arguments["spreadsheet_id"]
            comment_text = arguments["comment"]
            range_a1: Optional[str] = arguments.get("range")

            drive_service = self.auth.get_drive_service()
            body: Dict[str, Any] = {"content": comment_text}

            if range_a1:
                sheets_service = self.auth.get_sheets_service()
                body["anchor"] = _build_sheets_anchor(
                    sheets_service, spreadsheet_id, range_a1
                )

            result = (
                drive_service.comments()
                .create(fileId=spreadsheet_id, body=body, fields=_COMMENT_FIELDS)
                .execute()
            )

            summary = (
                f"Created comment {result.get('id')} on {range_a1}"
                if range_a1
                else f"Created file-level comment {result.get('id')}"
            )
            return self.format_success_response(
                json.dumps(result, indent=2, ensure_ascii=False),
                summary,
            )

        except HttpError as e:
            if e.resp.status == 404:
                raise SheetsAPIError("Spreadsheet not found", 404)
            raise DriveAPIError(f"Drive API error: {e.reason}", e.resp.status)
        except (InvalidRangeError, SheetsAPIError, DriveAPIError):
            raise
        except Exception as e:
            return self.format_error_response(e)


class ListCommentsHandler(SheetsToolHandler):
    """List threaded comments (and their replies) on a Google Sheet."""

    def __init__(self, auth: GoogleSheetsAuth) -> None:
        super().__init__(
            name="list_comments",
            description=(
                "List threaded comments on a Google Sheet, including replies, anchor "
                "(cell location), author, timestamps, and resolved status. Paginated; "
                "pass `page_token` from the previous response to fetch the next page."
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
                    "include_deleted": {
                        "type": "boolean",
                        "description": "Include deleted comments (default: false)",
                        "default": False,
                    },
                    "page_size": {
                        "type": "integer",
                        "description": "Max comments per page (1-100, default: 50)",
                        "default": 50,
                    },
                    "page_token": {
                        "type": "string",
                        "description": "Pagination token from a previous response",
                    },
                },
                "required": ["spreadsheet_id"],
            },
        )

    async def execute(self, arguments: Dict[str, Any]) -> List[TextContent]:
        try:
            self.validate_arguments(arguments, ["spreadsheet_id"])

            spreadsheet_id = arguments["spreadsheet_id"]
            include_deleted = bool(arguments.get("include_deleted", False))
            page_size = int(arguments.get("page_size", 50))
            page_token = arguments.get("page_token")

            drive_service = self.auth.get_drive_service()
            request_kwargs: Dict[str, Any] = {
                "fileId": spreadsheet_id,
                "fields": _COMMENT_LIST_FIELDS,
                "includeDeleted": include_deleted,
                "pageSize": max(1, min(page_size, 100)),
            }
            if page_token:
                request_kwargs["pageToken"] = page_token

            result = drive_service.comments().list(**request_kwargs).execute()

            comments = result.get("comments", [])
            response_data = {
                "comment_count": len(comments),
                "next_page_token": result.get("nextPageToken"),
                "comments": comments,
            }
            return self.format_success_response(
                json.dumps(response_data, indent=2, ensure_ascii=False),
                f"Found {len(comments)} comment(s)",
            )

        except HttpError as e:
            if e.resp.status == 404:
                raise SheetsAPIError("Spreadsheet not found", 404)
            raise DriveAPIError(f"Drive API error: {e.reason}", e.resp.status)
        except (SheetsAPIError, DriveAPIError):
            raise
        except Exception as e:
            return self.format_error_response(e)


class ReplyToCommentHandler(SheetsToolHandler):
    """Reply to an existing threaded comment, optionally resolving or reopening the thread."""

    def __init__(self, auth: GoogleSheetsAuth) -> None:
        super().__init__(
            name="reply_to_comment",
            description=(
                "Add a reply to an existing threaded comment on a Google Sheet. "
                "Use `action='resolve'` to mark the thread resolved, or `action='reopen'` "
                "to reopen a resolved thread. Either `reply` text or an `action` is required."
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
                    "comment_id": {
                        "type": "string",
                        "description": "The ID of the comment to reply to (from list_comments or add_comment)",
                    },
                    "reply": {
                        "type": "string",
                        "description": "Reply text. Optional if `action` is set, required otherwise.",
                    },
                    "action": {
                        "type": "string",
                        "description": "Optional thread action to apply with this reply",
                        "enum": ["resolve", "reopen"],
                    },
                },
                "required": ["spreadsheet_id", "comment_id"],
            },
        )

    async def execute(self, arguments: Dict[str, Any]) -> List[TextContent]:
        try:
            self.validate_arguments(arguments, ["spreadsheet_id", "comment_id"])

            spreadsheet_id = arguments["spreadsheet_id"]
            comment_id = arguments["comment_id"]
            reply_text: Optional[str] = arguments.get("reply")
            action: Optional[str] = arguments.get("action")

            if not reply_text and not action:
                raise ValueError("Provide `reply` text and/or an `action` (resolve/reopen).")

            body: Dict[str, Any] = {}
            if reply_text:
                body["content"] = reply_text
            if action:
                body["action"] = action
            else:
                # Drive requires non-empty content when no action is set; already enforced above.
                pass

            drive_service = self.auth.get_drive_service()
            result = (
                drive_service.replies()
                .create(
                    fileId=spreadsheet_id,
                    commentId=comment_id,
                    body=body,
                    fields=_REPLY_FIELDS,
                )
                .execute()
            )

            summary_parts = [f"Added reply {result.get('id')} to comment {comment_id}"]
            if action:
                summary_parts.append(f"(action: {action})")
            return self.format_success_response(
                json.dumps(result, indent=2, ensure_ascii=False),
                " ".join(summary_parts),
            )

        except HttpError as e:
            if e.resp.status == 404:
                raise SheetsAPIError(
                    "Spreadsheet or comment not found", 404
                )
            raise DriveAPIError(f"Drive API error: {e.reason}", e.resp.status)
        except (SheetsAPIError, DriveAPIError, ValueError):
            raise
        except Exception as e:
            return self.format_error_response(e)


COMMENTS_HANDLERS = [
    AddCommentHandler,
    ListCommentsHandler,
    ReplyToCommentHandler,
]
