"""MCP server exposing Microsoft Visio flowchart automation tools."""

from __future__ import annotations

from enum import Enum
from typing import Annotated, Any

from mcp.server.fastmcp import FastMCP
from pydantic import BaseModel, ConfigDict, Field

from .visio import DEFAULT_LINE_COLOR, VisioAutomationError, VisioController


mcp = FastMCP("visio_mcp")
controller = VisioController()


class ShapeKind(str, Enum):
    """Supported flowchart shape kinds."""

    PROCESS = "process"
    DECISION = "decision"
    START_END = "start_end"
    TERMINATOR = "terminator"
    DATA = "data"
    DOCUMENT = "document"
    CIRCLE = "circle"
    TEXT = "text"


class LayoutDirection(str, Enum):
    """Supported automatic layout directions."""

    TOP_TO_BOTTOM = "top_to_bottom"
    LEFT_TO_RIGHT = "left_to_right"


class FlowchartNode(BaseModel):
    """A node to draw in a flowchart."""

    model_config = ConfigDict(str_strip_whitespace=True, extra="forbid")

    id: str = Field(..., description="Stable node key used by edges, for example 'approve'.", min_length=1)
    text: str = Field(..., description="Text shown inside the Visio shape.", min_length=1)
    kind: ShapeKind = Field(default=ShapeKind.PROCESS, description="Flowchart shape kind.")
    width: float = Field(default=1.8, description="Shape width in inches.", gt=0.2, le=6.0)
    height: float = Field(default=0.8, description="Shape height in inches.", gt=0.2, le=4.0)
    fill_color: str | None = Field(default=None, description="Optional fill color in #RRGGBB format.")
    line_color: str = Field(default=DEFAULT_LINE_COLOR, description="Optional line color in #RRGGBB format.")


class FlowchartEdge(BaseModel):
    """A directed edge between two flowchart nodes."""

    model_config = ConfigDict(str_strip_whitespace=True, extra="forbid")

    from_: str = Field(..., alias="from", description="Source node ID.", min_length=1)
    to: str = Field(..., description="Target node ID.", min_length=1)
    text: str = Field(default="", description="Optional connector label, for example 'Yes' or 'No'.")
    line_color: str = Field(default=DEFAULT_LINE_COLOR, description="Optional connector color in #RRGGBB format.")


def _ok(result: dict[str, Any]) -> dict[str, Any]:
    """Wrap successful tool output consistently."""
    return {"ok": True, **result}


def _error(exc: Exception) -> dict[str, Any]:
    """Wrap tool errors with actionable messages."""
    if isinstance(exc, VisioAutomationError):
        return {"ok": False, "error": str(exc)}
    return {
        "ok": False,
        "error": f"Unexpected {type(exc).__name__}: {exc}. Check that Visio is installed and not showing a modal dialog.",
    }


@mcp.tool(
    name="visio_get_status",
    annotations={
        "title": "Get Visio automation status",
        "readOnlyHint": True,
        "destructiveHint": False,
        "idempotentHint": True,
        "openWorldHint": False,
    },
)
def visio_get_status() -> dict[str, Any]:
    """Get Microsoft Visio automation status without creating a new document.

    Returns:
        dict[str, Any]: JSON-serializable status with keys:
            - ok (bool): whether the tool call succeeded
            - connected (bool): whether a Visio COM application is available
            - visible (bool): Visio UI visibility when connected
            - version (str): Visio version when connected
            - documents (list[dict]): open document metadata
    """
    return _ok(controller.status())


@mcp.tool(
    name="visio_create_document",
    annotations={
        "title": "Create a Visio document",
        "readOnlyHint": False,
        "destructiveHint": False,
        "idempotentHint": False,
        "openWorldHint": False,
    },
)
def visio_create_document(
    visible: Annotated[bool, Field(description="Whether the Visio UI should be visible.")] = True,
    page_name: Annotated[
        str,
        Field(description="Name for the initial page.", min_length=1, max_length=100),
    ] = "Flowchart",
) -> dict[str, Any]:
    """Create a blank Visio document for drawing flowcharts.

    Args:
        visible (bool): Show or hide the Visio UI.
        page_name (str): Initial page name.

    Returns:
        dict[str, Any]: Document metadata including active page, path, page width, and page height.
    """
    try:
        return _ok(controller.create_document(visible=visible, page_name=page_name))
    except Exception as exc:
        return _error(exc)


@mcp.tool(
    name="visio_open_document",
    annotations={
        "title": "Open a Visio document",
        "readOnlyHint": False,
        "destructiveHint": False,
        "idempotentHint": False,
        "openWorldHint": False,
    },
)
def visio_open_document(
    path: Annotated[
        str,
        Field(description="Absolute or relative path to a .vsdx/.vsd Visio document.", min_length=1),
    ],
    visible: Annotated[bool, Field(description="Whether the Visio UI should be visible.")] = True,
) -> dict[str, Any]:
    """Open an existing Visio document and make it active.

    Args:
        path (str): Absolute or relative path to a .vsdx/.vsd file.
        visible (bool): Show or hide the Visio UI.

    Returns:
        dict[str, Any]: Document metadata for the opened file.
    """
    try:
        return _ok(controller.open_document(path=path, visible=visible))
    except Exception as exc:
        return _error(exc)


@mcp.tool(
    name="visio_save_document",
    annotations={
        "title": "Save the active Visio document",
        "readOnlyHint": False,
        "destructiveHint": False,
        "idempotentHint": True,
        "openWorldHint": False,
    },
)
def visio_save_document(
    path: Annotated[
        str | None,
        Field(description="Optional output path. If omitted, saves the active document to its existing location."),
    ] = None,
) -> dict[str, Any]:
    """Save the active Visio document.

    Args:
        path (str | None): Optional SaveAs destination. If omitted, saves to the current document path.

    Returns:
        dict[str, Any]: Updated document metadata.
    """
    try:
        return _ok(controller.save_document(path=path))
    except Exception as exc:
        return _error(exc)


@mcp.tool(
    name="visio_export_page",
    annotations={
        "title": "Export the active Visio page",
        "readOnlyHint": False,
        "destructiveHint": False,
        "idempotentHint": True,
        "openWorldHint": False,
    },
)
def visio_export_page(
    path: Annotated[
        str,
        Field(description="Export path ending in pdf, png, svg, jpg, bmp, gif, tif, or emf.", min_length=1),
    ],
) -> dict[str, Any]:
    """Export the active Visio page to PDF, PNG, SVG, JPG, BMP, GIF, TIF, or EMF.

    Args:
        path (str): Output file path with an extension supported by Visio Page.Export.

    Returns:
        dict[str, Any]: Export status and resolved output path.
    """
    try:
        return _ok(controller.export_page(path=path))
    except Exception as exc:
        return _error(exc)


@mcp.tool(
    name="visio_add_shape",
    annotations={
        "title": "Add a Visio flowchart shape",
        "readOnlyHint": False,
        "destructiveHint": False,
        "idempotentHint": False,
        "openWorldHint": False,
    },
)
def visio_add_shape(
    text: Annotated[str, Field(description="Shape label text.", min_length=1)],
    x: Annotated[float, Field(description="Shape center X coordinate in inches.")],
    y: Annotated[float, Field(description="Shape center Y coordinate in inches.")],
    kind: Annotated[ShapeKind, Field(description="Shape kind to draw.")] = ShapeKind.PROCESS,
    width: Annotated[float, Field(description="Shape width in inches.", gt=0.2, le=6.0)] = 1.8,
    height: Annotated[float, Field(description="Shape height in inches.", gt=0.2, le=4.0)] = 0.8,
    fill_color: Annotated[str | None, Field(description="Optional fill color in #RRGGBB format.")] = None,
    line_color: Annotated[str, Field(description="Optional line color in #RRGGBB format.")] = DEFAULT_LINE_COLOR,
    text_color: Annotated[str, Field(description="Optional text color in #RRGGBB format.")] = "#1F2933",
    name: Annotated[str | None, Field(description="Optional Visio universal shape name.")] = None,
) -> dict[str, Any]:
    """Add a single flowchart shape to the active Visio page.

    Args:
        text (str): Label text.
        x/y/width/height (float): Coordinates and size in inches.
        kind (ShapeKind): process, decision, start_end, data, document, circle, or text.
        fill_color/line_color/text_color (str): Optional #RRGGBB colors.
        name (str | None): Optional universal shape name.

    Returns:
        dict[str, Any]: Created shape metadata including Visio shape ID.
    """
    try:
        shape = controller.add_shape(
            kind=kind.value,
            text=text,
            x=x,
            y=y,
            width=width,
            height=height,
            fill_color=fill_color,
            line_color=line_color,
            text_color=text_color,
            name=name,
        )
        return _ok({"shape": shape.to_dict()})
    except Exception as exc:
        return _error(exc)


@mcp.tool(
    name="visio_connect_shapes",
    annotations={
        "title": "Connect two Visio shapes",
        "readOnlyHint": False,
        "destructiveHint": False,
        "idempotentHint": False,
        "openWorldHint": False,
    },
)
def visio_connect_shapes(
    from_shape_id: Annotated[int, Field(description="Source Visio shape ID returned by visio_add_shape.", ge=1)],
    to_shape_id: Annotated[int, Field(description="Target Visio shape ID returned by visio_add_shape.", ge=1)],
    text: Annotated[str, Field(description="Optional connector label.")] = "",
    line_color: Annotated[str, Field(description="Optional line color in #RRGGBB format.")] = DEFAULT_LINE_COLOR,
) -> dict[str, Any]:
    """Draw an arrow connector between two existing Visio shapes.

    Args:
        from_shape_id (int): Source shape ID.
        to_shape_id (int): Target shape ID.
        text (str): Optional label.
        line_color (str): Optional #RRGGBB connector color.

    Returns:
        dict[str, Any]: Created connector metadata including Visio shape ID.
    """
    try:
        connector = controller.connect_shapes(
            from_shape_id=from_shape_id,
            to_shape_id=to_shape_id,
            text=text,
            line_color=line_color,
        )
        return _ok({"connector": connector.to_dict()})
    except Exception as exc:
        return _error(exc)


@mcp.tool(
    name="visio_draw_flowchart",
    annotations={
        "title": "Draw a complete flowchart",
        "readOnlyHint": False,
        "destructiveHint": False,
        "idempotentHint": False,
        "openWorldHint": False,
    },
)
def visio_draw_flowchart(
    nodes: Annotated[list[FlowchartNode], Field(description="Nodes to draw.", min_length=1, max_length=100)],
    edges: Annotated[
        list[FlowchartEdge] | None,
        Field(description="Directed edges between node IDs.", max_length=200),
    ] = None,
    direction: Annotated[LayoutDirection, Field(description="Automatic layout direction.")] = LayoutDirection.TOP_TO_BOTTOM,
    page_name: Annotated[str | None, Field(description="Optional active page name.", max_length=100)] = "Flowchart",
    clear_page: Annotated[bool, Field(description="Delete existing shapes on the active page before drawing.")] = True,
    auto_size_page: Annotated[bool, Field(description="Expand page size to fit the generated layout.")] = True,
) -> dict[str, Any]:
    """Draw a complete Visio flowchart from nodes and directed edges.

    This is the preferred high-level workflow tool for creating a process diagram in one call.
    It creates a Visio document if none is open, optionally clears the active page, performs
    a deterministic layered layout, draws all shapes, and connects them with arrow lines.

    Args:
        nodes (list[FlowchartNode]): Shape IDs, text, kind, size, and optional colors.
        edges (list[FlowchartEdge] | None): Directed edges using node IDs and optional labels.
        direction (LayoutDirection): top_to_bottom or left_to_right.
        page_name (str | None): Optional active page name.
        clear_page (bool): Whether to delete current page shapes first.
        auto_size_page (bool): Whether to expand the page to fit the generated layout.

    Returns:
        dict[str, Any]: Document metadata, node IDs mapped to Visio shape IDs, and connector metadata.
    """
    try:
        return _ok(
            controller.draw_flowchart(
                nodes=[node.model_dump() for node in nodes],
                edges=[edge.model_dump(by_alias=True) for edge in edges or []],
                direction=direction.value,
                page_name=page_name,
                clear_page=clear_page,
                auto_size_page=auto_size_page,
            )
        )
    except Exception as exc:
        return _error(exc)


@mcp.tool(
    name="visio_clear_page",
    annotations={
        "title": "Clear the active Visio page",
        "readOnlyHint": False,
        "destructiveHint": True,
        "idempotentHint": True,
        "openWorldHint": False,
    },
)
def visio_clear_page() -> dict[str, Any]:
    """Delete all shapes from the active Visio page.

    Returns:
        dict[str, Any]: Number of deleted shapes and active page name.
    """
    try:
        return _ok(controller.clear_page())
    except Exception as exc:
        return _error(exc)


@mcp.tool(
    name="visio_list_shapes",
    annotations={
        "title": "List active page shapes",
        "readOnlyHint": True,
        "destructiveHint": False,
        "idempotentHint": True,
        "openWorldHint": False,
    },
)
def visio_list_shapes() -> dict[str, Any]:
    """List shapes currently present on the active Visio page.

    Returns:
        dict[str, Any]: Shape metadata including IDs, names, text, position, size, and 1D status.
    """
    try:
        return _ok(controller.list_shapes())
    except Exception as exc:
        return _error(exc)


def main() -> None:
    """Run the Visio MCP server over stdio."""
    mcp.run()


if __name__ == "__main__":
    main()
