"""Main MCP server for Claude Google Sheets integration."""

import argparse
import asyncio
import logging
import sys
from typing import Any, Sequence

from mcp.server import Server
from mcp.types import (
    CallToolRequest,
    ListToolsRequest,
    TextContent,
    Tool,
)

from .auth.oauth_manager import GoogleSheetsAuth
from .core.exceptions import AuthenticationError, GoogleSheetsMCPError
from .tools.comments_tools import COMMENTS_HANDLERS
from .tools.drive_tools import DRIVE_HANDLERS
from .tools.sheets_tools import SHEETS_HANDLERS

# Setup logging
logging.basicConfig(
    level=logging.INFO, format="%(asctime)s - %(name)s - %(levelname)s - %(message)s"
)
logger = logging.getLogger(__name__)

# Create MCP server instance
app = Server("claude-google-sheets-mcp")

# Global variables
auth_manager: GoogleSheetsAuth = None
tool_handlers = {}


def initialize_handlers(auth: GoogleSheetsAuth) -> None:
    """Initialize all tool handlers with authentication."""
    global tool_handlers

    # Initialize drive handlers
    for handler_class in DRIVE_HANDLERS:
        handler = handler_class(auth)
        tool_handlers[handler.name] = handler
        logger.info(f"Registered drive tool: {handler.name}")

    # Initialize sheets handlers
    for handler_class in SHEETS_HANDLERS:
        handler = handler_class(auth)
        tool_handlers[handler.name] = handler
        logger.info(f"Registered sheets tool: {handler.name}")

    # Initialize comments handlers (threaded Drive comments anchored to sheet cells)
    for handler_class in COMMENTS_HANDLERS:
        handler = handler_class(auth)
        tool_handlers[handler.name] = handler
        logger.info(f"Registered comments tool: {handler.name}")


@app.list_tools()
async def list_tools() -> list[Tool]:
    """List all available tools."""
    tools = []
    for handler in tool_handlers.values():
        tools.append(handler.get_tool_definition())

    logger.info(f"Listed {len(tools)} available tools")
    return tools


@app.call_tool()
async def call_tool(name: str, arguments: dict[str, Any]) -> Sequence[TextContent]:
    """Execute a tool with the given arguments."""
    try:
        logger.info(f"Executing tool: {name} with arguments: {arguments}")

        if name not in tool_handlers:
            error_msg = f"Unknown tool: {name}"
            logger.error(error_msg)
            return [TextContent(type="text", text=error_msg)]

        handler = tool_handlers[name]
        result = await handler.execute(arguments)

        logger.info(f"Successfully executed tool: {name}")
        return result

    except GoogleSheetsMCPError as e:
        error_msg = f"Tool execution error: {e.message}"
        if e.details:
            error_msg += f" Details: {e.details}"
        logger.error(error_msg)
        return [TextContent(type="text", text=error_msg)]

    except Exception as e:
        error_msg = f"Unexpected error executing tool {name}: {str(e)}"
        logger.error(error_msg)
        return [TextContent(type="text", text=error_msg)]


def setup_auth(credentials_dir: str = None) -> GoogleSheetsAuth:
    """Set up authentication manager."""
    try:
        auth = GoogleSheetsAuth(credentials_dir)
        auth.authenticate()

        # Check account type and permissions
        try:
            account_type = auth.detect_account_type()
            logger.info(f"Detected account type: {account_type}")

            permissions = auth.check_permissions()
            available_perms = [k for k, v in permissions.items() if v]
            if available_perms:
                logger.info(f"Available permissions: {', '.join(available_perms)}")
            else:
                logger.warning("No API permissions detected - functionality may be limited")

            # Try to get user info for logging, but don't fail if not available
            try:
                user_info = auth.get_user_info()
                logger.info(
                    f"Authenticated as: {user_info.get('email', 'Unknown user')}"
                )
            except Exception:
                logger.info("Authentication successful (user info not available)")

        except Exception as e:
            logger.debug(f"Account detection failed: {str(e)}")
            logger.info("Authentication successful")

        return auth
    except Exception as e:
        logger.error(f"Authentication failed: {str(e)}")
        raise AuthenticationError(f"Failed to authenticate: {str(e)}") from e


async def main() -> None:
    """Main entry point for the MCP server."""
    parser = argparse.ArgumentParser(description="Claude Google Sheets MCP Server")
    parser.add_argument(
        "--credentials-dir",
        type=str,
        help="Directory containing Google API credentials",
    )
    parser.add_argument("--debug", action="store_true", help="Enable debug logging")
    parser.add_argument(
        "--setup",
        action="store_true",
        help="Run interactive setup wizard",
    )

    args = parser.parse_args()

    # Run setup wizard if requested
    if args.setup:
        from .setup import SetupWizard

        wizard = SetupWizard(args.credentials_dir)
        success = wizard.run()
        sys.exit(0 if success else 1)

    if args.debug:
        logging.getLogger().setLevel(logging.DEBUG)
        logger.debug("Debug logging enabled")

    try:
        # Initialize authentication
        global auth_manager
        auth_manager = setup_auth(args.credentials_dir)

        # Initialize tool handlers
        initialize_handlers(auth_manager)

        logger.info("Claude Google Sheets MCP Server started successfully")
        logger.info(f"Registered {len(tool_handlers)} tool handlers")

        # Run the server
        from mcp.server.stdio import stdio_server

        async with stdio_server() as (read_stream, write_stream):
            await app.run(
                read_stream, write_stream, app.create_initialization_options()
            )

    except AuthenticationError as e:
        logger.error(f"Authentication error: {e.message}")
        sys.exit(1)
    except Exception as e:
        logger.error(f"Server startup failed: {str(e)}")
        sys.exit(1)


def cli_main() -> None:
    """CLI entry point."""
    try:
        asyncio.run(main())
    except KeyboardInterrupt:
        logger.info("Server shutdown requested")
    except Exception as e:
        logger.error(f"Server error: {str(e)}")
        sys.exit(1)


if __name__ == "__main__":
    cli_main()
