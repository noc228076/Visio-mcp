"""Microsoft Visio COM automation helpers."""

from __future__ import annotations

import math
import os
import re
import threading
from array import array
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Literal


class VisioAutomationError(RuntimeError):
    """Raised when Visio cannot complete an automation request."""


ShapeKind = Literal["process", "decision", "start_end", "terminator", "data", "document", "circle", "text"]
Direction = Literal["top_to_bottom", "left_to_right"]


DEFAULT_FILL_COLORS: dict[str, str] = {
    "process": "#EAF2FF",
    "decision": "#FFF4D6",
    "start_end": "#E7F8EF",
    "terminator": "#E7F8EF",
    "data": "#F2EBFF",
    "document": "#F7F7F7",
    "circle": "#FEECEC",
    "text": "#FFFFFF",
}
DEFAULT_LINE_COLOR = "#335C81"
DEFAULT_TEXT_COLOR = "#1F2933"
PAGE_WIDTH = 11.0
PAGE_HEIGHT = 8.5
VISIO_POLYLINE_1D = 8


@dataclass(frozen=True)
class DrawnShape:
    """Metadata for a shape created on the active Visio page."""

    id: int
    name: str
    kind: str
    text: str
    pin_x: float
    pin_y: float
    width: float
    height: float

    def to_dict(self) -> dict[str, Any]:
        """Return JSON-serializable metadata."""
        return {
            "id": self.id,
            "name": self.name,
            "kind": self.kind,
            "text": self.text,
            "pin_x": self.pin_x,
            "pin_y": self.pin_y,
            "width": self.width,
            "height": self.height,
        }


@dataclass(frozen=True)
class DrawnConnector:
    """Metadata for a connector created on the active Visio page."""

    id: int
    name: str
    from_shape_id: int
    to_shape_id: int
    text: str

    def to_dict(self) -> dict[str, Any]:
        """Return JSON-serializable metadata."""
        return {
            "id": self.id,
            "name": self.name,
            "from_shape_id": self.from_shape_id,
            "to_shape_id": self.to_shape_id,
            "text": self.text,
        }


def _import_visio_dependencies() -> tuple[Any, Any, Any]:
    """Import pywin32 modules lazily so the package can still be inspected elsewhere."""
    if os.name != "nt":
        raise VisioAutomationError("Microsoft Visio automation requires Windows and pywin32.")

    try:
        import pythoncom  # type: ignore[import-not-found]
        import pywintypes  # type: ignore[import-not-found]
        import win32com.client  # type: ignore[import-not-found]
    except ImportError as exc:
        raise VisioAutomationError(
            "pywin32 is required. Install dependencies with `pip install -e .` on Windows."
        ) from exc

    return pythoncom, pywintypes, win32com.client


def _rgb_formula(hex_color: str) -> str:
    """Convert #RRGGBB into a Visio RGB formula."""
    value = hex_color.strip()
    if not re.fullmatch(r"#?[0-9a-fA-F]{6}", value):
        raise VisioAutomationError(f"Invalid color '{hex_color}'. Use #RRGGBB, for example #EAF2FF.")
    value = value.removeprefix("#")
    red = int(value[0:2], 16)
    green = int(value[2:4], 16)
    blue = int(value[4:6], 16)
    return f"RGB({red},{green},{blue})"


def _safe_name(text: str, fallback: str) -> str:
    """Create a Visio-friendly universal shape name."""
    base = re.sub(r"[^0-9A-Za-z_]+", "_", text.strip())[:40].strip("_")
    return base or fallback


def _resolve_output_path(path: str) -> str:
    """Resolve user supplied paths while preserving explicit absolute locations."""
    if not path.strip():
        raise VisioAutomationError("Path cannot be empty.")
    return str(Path(path).expanduser().resolve())


def _cell_set(shape: Any, cell_name: str, formula: str) -> None:
    """Set a Visio shape cell and ignore cells unavailable for a given shape type."""
    try:
        shape.CellsU(cell_name).FormulaU = formula
    except Exception:
        return


def _point_array(points: list[float]) -> Any:
    """Create a COM-friendly double array for Visio polyline drawing."""
    return array("d", points)


def _shape_text(shape: Any) -> str:
    """Read shape text safely."""
    try:
        return str(shape.Text)
    except Exception:
        return ""


class VisioController:
    """Stateful controller for a Microsoft Visio COM automation session."""

    def __init__(self) -> None:
        self._app: Any | None = None
        self._app_thread_id: int | None = None
        self._pythoncom: Any | None = None
        self._pywintypes: Any | None = None
        self._win32com_client: Any | None = None

    def status(self) -> dict[str, Any]:
        """Return status for the current Visio automation session."""
        try:
            app = self._get_app(create_if_missing=False)
        except VisioAutomationError as exc:
            return {"connected": False, "message": str(exc)}

        documents = []
        try:
            for index in range(1, app.Documents.Count + 1):
                document = app.Documents.Item(index)
                documents.append(
                    {
                        "name": str(document.Name),
                        "path": str(document.FullName) if document.Path else "",
                        "saved": bool(document.Saved),
                    }
                )
        except Exception:
            documents = []

        return {
            "connected": True,
            "visible": bool(app.Visible),
            "version": str(app.Version),
            "document_count": len(documents),
            "documents": documents,
        }

    def create_document(self, *, visible: bool = True, page_name: str = "Flowchart") -> dict[str, Any]:
        """Create a blank Visio document and prepare one flowchart page."""
        app = self._get_app(create_if_missing=True)
        app.Visible = visible
        document = app.Documents.Add("")
        page = app.ActivePage
        if page_name:
            page.Name = page_name
        self._set_page_size(page, PAGE_WIDTH, PAGE_HEIGHT)
        return self._document_info(document=document, page=page)

    def open_document(self, *, path: str, visible: bool = True) -> dict[str, Any]:
        """Open an existing Visio document."""
        app = self._get_app(create_if_missing=True)
        app.Visible = visible
        resolved = _resolve_output_path(path)
        if not Path(resolved).exists():
            raise VisioAutomationError(f"Visio document not found: {resolved}")
        document = app.Documents.Open(resolved)
        return self._document_info(document=document, page=app.ActivePage)

    def save_document(self, *, path: str | None = None) -> dict[str, Any]:
        """Save the active Visio document, optionally using SaveAs for a new path."""
        document = self._active_document()
        if path:
            resolved = _resolve_output_path(path)
            Path(resolved).parent.mkdir(parents=True, exist_ok=True)
            document.SaveAs(resolved)
        else:
            if not document.Path:
                raise VisioAutomationError("The active document has no path. Provide a `path` to save it.")
            document.Save()
        return self._document_info(document=document, page=self._active_page())

    def export_page(self, *, path: str) -> dict[str, Any]:
        """Export the active Visio page to a file supported by Visio, such as PDF, PNG, SVG, or JPG."""
        page = self._active_page()
        resolved = _resolve_output_path(path)
        suffix = Path(resolved).suffix.lower()
        if suffix not in {".pdf", ".png", ".svg", ".jpg", ".jpeg", ".bmp", ".gif", ".tif", ".tiff", ".emf"}:
            raise VisioAutomationError(
                "Unsupported export extension. Use one of: pdf, png, svg, jpg, bmp, gif, tif, emf."
            )
        Path(resolved).parent.mkdir(parents=True, exist_ok=True)
        if suffix == ".pdf":
            document = self._active_document()
            document.ExportAsFixedFormat(1, resolved, 1, 0)
        else:
            page.Export(resolved)
        return {"exported": True, "path": resolved, "page": str(page.Name)}

    def add_shape(
        self,
        *,
        kind: ShapeKind,
        text: str,
        x: float,
        y: float,
        width: float = 1.7,
        height: float = 0.75,
        fill_color: str | None = None,
        line_color: str = DEFAULT_LINE_COLOR,
        text_color: str = DEFAULT_TEXT_COLOR,
        name: str | None = None,
    ) -> DrawnShape:
        """Add one flowchart shape to the active page."""
        kind_value = str(getattr(kind, "value", kind))
        page = self._active_page()
        shape = self._draw_shape(
            page=page,
            kind=kind_value,
            x=x,
            y=y,
            width=width,
            height=height,
        )
        shape.Text = text
        self._style_shape(
            shape,
            kind=kind_value,
            fill_color=fill_color or DEFAULT_FILL_COLORS.get(kind_value, "#FFFFFF"),
            line_color=line_color,
            text_color=text_color,
        )
        shape.NameU = self._unique_name(page, name or _safe_name(text, kind_value))
        return self._drawn_shape(shape=shape, kind=kind_value, text=text)

    def connect_shapes(
        self,
        *,
        from_shape_id: int,
        to_shape_id: int,
        text: str = "",
        line_color: str = DEFAULT_LINE_COLOR,
    ) -> DrawnConnector:
        """Draw an arrow connector between two existing shapes on the active page."""
        page = self._active_page()
        source = page.Shapes.ItemFromID(from_shape_id)
        target = page.Shapes.ItemFromID(to_shape_id)
        source_point, target_point = self._edge_points(source, target)
        connector = page.DrawLine(source_point[0], source_point[1], target_point[0], target_point[1])
        connector.Text = text
        _cell_set(connector, "LineColor", _rgb_formula(line_color))
        _cell_set(connector, "LineWeight", "1.25 pt")
        _cell_set(connector, "EndArrow", "13")
        connector.NameU = self._unique_name(page, f"connector_{from_shape_id}_to_{to_shape_id}")
        return DrawnConnector(
            id=int(connector.ID),
            name=str(connector.NameU),
            from_shape_id=from_shape_id,
            to_shape_id=to_shape_id,
            text=text,
        )

    def draw_flowchart(
        self,
        *,
        nodes: list[dict[str, Any]],
        edges: list[dict[str, Any]],
        direction: Direction = "top_to_bottom",
        page_name: str | None = None,
        clear_page: bool = True,
        auto_size_page: bool = True,
    ) -> dict[str, Any]:
        """Draw a full flowchart from node and edge specifications."""
        app = self._get_app(create_if_missing=True)
        if app.Documents.Count == 0:
            self.create_document(visible=True, page_name=page_name or "Flowchart")
        page = self._active_page()
        if page_name:
            page.Name = page_name
        if clear_page:
            self.clear_page()

        layout = self._layout_nodes(nodes, edges, direction=direction)
        if auto_size_page:
            self._fit_page_to_layout(page, layout)

        created_by_key: dict[str, DrawnShape] = {}
        for node in nodes:
            key = str(node["id"])
            x, y = layout[key]
            created = self.add_shape(
                kind=node.get("kind", "process"),
                text=str(node.get("text", key)),
                x=x,
                y=y,
                width=float(node.get("width", 1.8)),
                height=float(node.get("height", 0.8)),
                fill_color=node.get("fill_color"),
                line_color=node.get("line_color", DEFAULT_LINE_COLOR),
                name=_safe_name(key, "node"),
            )
            created_by_key[key] = created

        connectors: list[DrawnConnector] = []
        for edge in edges:
            source_key = str(edge["from"])
            target_key = str(edge["to"])
            if source_key not in created_by_key or target_key not in created_by_key:
                raise VisioAutomationError(
                    f"Edge references unknown node: {source_key!r} -> {target_key!r}."
                )
            connectors.append(
                self.connect_shapes(
                    from_shape_id=created_by_key[source_key].id,
                    to_shape_id=created_by_key[target_key].id,
                    text=str(edge.get("text", "")),
                    line_color=str(edge.get("line_color", DEFAULT_LINE_COLOR)),
                )
            )

        return {
            "document": self._document_info(document=self._active_document(), page=page),
            "direction": direction,
            "node_count": len(created_by_key),
            "edge_count": len(connectors),
            "nodes": {key: value.to_dict() for key, value in created_by_key.items()},
            "connectors": [connector.to_dict() for connector in connectors],
        }

    def clear_page(self) -> dict[str, Any]:
        """Delete all shapes from the active Visio page."""
        page = self._active_page()
        deleted = 0
        for index in range(page.Shapes.Count, 0, -1):
            page.Shapes.Item(index).Delete()
            deleted += 1
        return {"cleared": True, "deleted_shapes": deleted, "page": str(page.Name)}

    def list_shapes(self) -> dict[str, Any]:
        """List shapes on the active page."""
        page = self._active_page()
        shapes: list[dict[str, Any]] = []
        for index in range(1, page.Shapes.Count + 1):
            shape = page.Shapes.Item(index)
            shapes.append(
                {
                    "id": int(shape.ID),
                    "name": str(shape.NameU),
                    "text": _shape_text(shape),
                    "pin_x": round(float(shape.CellsU("PinX").ResultIU), 4),
                    "pin_y": round(float(shape.CellsU("PinY").ResultIU), 4),
                    "width": round(float(shape.CellsU("Width").ResultIU), 4),
                    "height": round(float(shape.CellsU("Height").ResultIU), 4),
                    "one_d": bool(shape.OneD),
                }
            )
        return {
            "page": str(page.Name),
            "count": len(shapes),
            "shapes": shapes,
        }

    def _get_app(self, *, create_if_missing: bool) -> Any:
        """Return an active Visio application object."""
        current_thread_id = threading.get_ident()
        if self._app is not None and self._app_thread_id == current_thread_id:
            try:
                _ = self._app.Version
                return self._app
            except Exception:
                self._app = None
                self._app_thread_id = None
        elif self._app is not None:
            self._app = None
            self._app_thread_id = None

        pythoncom, pywintypes, win32com_client = _import_visio_dependencies()
        self._pythoncom = pythoncom
        self._pywintypes = pywintypes
        self._win32com_client = win32com_client
        pythoncom.CoInitialize()

        try:
            self._app = win32com_client.GetActiveObject("Visio.Application")
        except pywintypes.com_error as exc:
            if not create_if_missing:
                raise VisioAutomationError(
                    "Visio is not running. Call visio_create_document or open Visio first."
                ) from exc
            try:
                self._app = win32com_client.Dispatch("Visio.Application")
            except pywintypes.com_error as dispatch_exc:
                raise VisioAutomationError(
                    "Unable to start Microsoft Visio through COM. Confirm Visio is installed and licensed."
                ) from dispatch_exc

        self._app_thread_id = current_thread_id
        return self._app

    def _active_document(self) -> Any:
        """Return the active Visio document."""
        app = self._get_app(create_if_missing=False)
        if app.Documents.Count == 0:
            raise VisioAutomationError("No Visio document is open. Create or open a document first.")
        return app.ActiveDocument

    def _active_page(self) -> Any:
        """Return the active Visio page."""
        app = self._get_app(create_if_missing=False)
        if app.Documents.Count == 0:
            raise VisioAutomationError("No Visio document is open. Create or open a document first.")
        return app.ActivePage

    def _document_info(self, *, document: Any, page: Any) -> dict[str, Any]:
        """Return document metadata."""
        return {
            "document_name": str(document.Name),
            "document_path": str(document.FullName) if document.Path else "",
            "saved": bool(document.Saved),
            "active_page": str(page.Name),
            "page_width": round(float(page.PageSheet.CellsU("PageWidth").ResultIU), 4),
            "page_height": round(float(page.PageSheet.CellsU("PageHeight").ResultIU), 4),
        }

    def _set_page_size(self, page: Any, width: float, height: float) -> None:
        """Set page size in Visio internal units, which are inches."""
        page.PageSheet.CellsU("PageWidth").FormulaU = f"{width} in"
        page.PageSheet.CellsU("PageHeight").FormulaU = f"{height} in"

    def _draw_shape(
        self,
        *,
        page: Any,
        kind: ShapeKind,
        x: float,
        y: float,
        width: float,
        height: float,
    ) -> Any:
        """Draw a flowchart shape using direct page drawing APIs."""
        left = x - width / 2
        right = x + width / 2
        bottom = y - height / 2
        top = y + height / 2

        if kind in {"process", "start_end", "terminator", "data", "document", "text"}:
            if kind == "data":
                skew = min(width * 0.18, 0.35)
                points = _point_array(
                    [
                        left + skew,
                        bottom,
                        right,
                        bottom,
                        right - skew,
                        top,
                        left,
                        top,
                        left + skew,
                        bottom,
                    ]
                )
                return page.DrawPolyline(points, 0)

            shape = page.DrawRectangle(left, bottom, right, top)
            if kind in {"start_end", "terminator"}:
                _cell_set(shape, "LineRounding", f"{min(width, height) / 2} in")
            elif kind == "document":
                self._add_document_wave(page, x=x, y=y, width=width, height=height)
            elif kind == "text":
                _cell_set(shape, "LinePattern", "0")
                _cell_set(shape, "FillPattern", "0")
            return shape

        if kind == "decision":
            points = _point_array([x, top, right, y, x, bottom, left, y, x, top])
            return page.DrawPolyline(points, 0)

        if kind == "circle":
            return page.DrawOval(left, bottom, right, top)

        raise VisioAutomationError(f"Unsupported shape kind: {kind}")

    def _add_document_wave(self, page: Any, *, x: float, y: float, width: float, height: float) -> None:
        """Add a subtle document bottom wave as a decorative line."""
        left = x - width / 2
        right = x + width / 2
        bottom = y - height / 2
        points = _point_array(
            [
                left,
                bottom + height * 0.18,
                left + width * 0.33,
                bottom + height * 0.04,
                left + width * 0.66,
                bottom + height * 0.24,
                right,
                bottom + height * 0.12,
            ]
        )
        try:
            wave = page.DrawPolyline(points, VISIO_POLYLINE_1D)
            _cell_set(wave, "LineColor", _rgb_formula("#8899AA"))
            _cell_set(wave, "LineWeight", "0.75 pt")
        except Exception:
            return

    def _style_shape(
        self,
        shape: Any,
        *,
        kind: str,
        fill_color: str,
        line_color: str,
        text_color: str,
    ) -> None:
        """Apply basic publication-friendly styling."""
        if kind != "text":
            _cell_set(shape, "FillForegnd", _rgb_formula(fill_color))
            _cell_set(shape, "FillPattern", "1")
            _cell_set(shape, "LineColor", _rgb_formula(line_color))
            _cell_set(shape, "LineWeight", "1.25 pt")
        _cell_set(shape, "Char.Color", _rgb_formula(text_color))
        _cell_set(shape, "Char.Font", 'FONT("Arial")')
        _cell_set(shape, "Char.Size", "8.5 pt")
        _cell_set(shape, "Para.HorzAlign", "1")
        _cell_set(shape, "VerticalAlign", "1")

    def _unique_name(self, page: Any, base_name: str) -> str:
        """Create a unique NameU value on a page."""
        existing = {str(page.Shapes.Item(index).NameU).lower() for index in range(1, page.Shapes.Count + 1)}
        base = _safe_name(base_name, "shape")
        candidate = base
        suffix = 2
        while candidate.lower() in existing:
            candidate = f"{base}_{suffix}"
            suffix += 1
        return candidate

    def _drawn_shape(self, *, shape: Any, kind: str, text: str) -> DrawnShape:
        """Build shape metadata from a COM shape object."""
        return DrawnShape(
            id=int(shape.ID),
            name=str(shape.NameU),
            kind=kind,
            text=text,
            pin_x=round(float(shape.CellsU("PinX").ResultIU), 4),
            pin_y=round(float(shape.CellsU("PinY").ResultIU), 4),
            width=round(float(shape.CellsU("Width").ResultIU), 4),
            height=round(float(shape.CellsU("Height").ResultIU), 4),
        )

    def _edge_points(self, source: Any, target: Any) -> tuple[tuple[float, float], tuple[float, float]]:
        """Pick approximate connection points on the boundary of two shapes."""
        sx = float(source.CellsU("PinX").ResultIU)
        sy = float(source.CellsU("PinY").ResultIU)
        sw = float(source.CellsU("Width").ResultIU)
        sh = float(source.CellsU("Height").ResultIU)
        tx = float(target.CellsU("PinX").ResultIU)
        ty = float(target.CellsU("PinY").ResultIU)
        tw = float(target.CellsU("Width").ResultIU)
        th = float(target.CellsU("Height").ResultIU)

        dx = tx - sx
        dy = ty - sy
        if abs(dx) > abs(dy):
            source_x = sx + math.copysign(sw / 2, dx)
            target_x = tx - math.copysign(tw / 2, dx)
            return (source_x, sy), (target_x, ty)

        source_y = sy + math.copysign(sh / 2, dy)
        target_y = ty - math.copysign(th / 2, dy)
        return (sx, source_y), (tx, target_y)

    def _layout_nodes(
        self,
        nodes: list[dict[str, Any]],
        edges: list[dict[str, Any]],
        *,
        direction: Direction,
    ) -> dict[str, tuple[float, float]]:
        """Create a deterministic layered layout for flowchart nodes."""
        if not nodes:
            raise VisioAutomationError("At least one node is required.")

        node_ids = [str(node["id"]) for node in nodes]
        if len(set(node_ids)) != len(node_ids):
            raise VisioAutomationError("Node IDs must be unique.")

        node_set = set(node_ids)
        adjacency: dict[str, list[str]] = {node_id: [] for node_id in node_ids}
        indegree: dict[str, int] = {node_id: 0 for node_id in node_ids}
        for edge in edges:
            source = str(edge["from"])
            target = str(edge["to"])
            if source not in node_set or target not in node_set:
                raise VisioAutomationError(f"Edge references unknown node: {source!r} -> {target!r}.")
            adjacency[source].append(target)
            indegree[target] += 1

        levels: dict[str, int] = {node_id: 0 for node_id in node_ids}
        queue = [node_id for node_id in node_ids if indegree[node_id] == 0]
        if not queue:
            queue = [node_ids[0]]
        seen = set(queue)

        while queue:
            current = queue.pop(0)
            for target in adjacency[current]:
                levels[target] = max(levels[target], levels[current] + 1)
                indegree[target] -= 1
                if indegree[target] == 0 and target not in seen:
                    queue.append(target)
                    seen.add(target)

        for node_id in node_ids:
            if node_id not in seen:
                levels[node_id] = max(levels.values(), default=0) + 1

        grouped: dict[int, list[str]] = {}
        for node_id in node_ids:
            grouped.setdefault(levels[node_id], []).append(node_id)

        horizontal_gap = 2.35
        vertical_gap = 1.45
        margin = 1.0
        max_level = max(grouped, default=0)
        max_group_size = max((len(group) for group in grouped.values()), default=1)
        required_width = max(PAGE_WIDTH, (max_group_size - 1) * horizontal_gap + margin * 2 + 2.0)
        required_height = max(PAGE_HEIGHT, max_level * vertical_gap + margin * 2 + 1.25)
        layout: dict[str, tuple[float, float]] = {}
        if direction == "top_to_bottom":
            for level in sorted(grouped):
                row = grouped[level]
                row_width = (len(row) - 1) * horizontal_gap
                for index, node_id in enumerate(row):
                    x = required_width / 2 - row_width / 2 + index * horizontal_gap
                    y = required_height - margin - level * vertical_gap
                    layout[node_id] = (round(x, 4), round(y, 4))
        else:
            for level in sorted(grouped):
                column = grouped[level]
                column_height = (len(column) - 1) * vertical_gap
                for index, node_id in enumerate(column):
                    x = margin + level * horizontal_gap
                    y = required_height / 2 + column_height / 2 - index * vertical_gap
                    layout[node_id] = (round(x, 4), round(y, 4))
        return layout

    def _fit_page_to_layout(self, page: Any, layout: dict[str, tuple[float, float]]) -> None:
        """Resize page if the generated layout needs more space."""
        max_x = max(x for x, _ in layout.values()) + 1.5
        max_y = max(y for _, y in layout.values()) + 1.0
        current_width = float(page.PageSheet.CellsU("PageWidth").ResultIU)
        current_height = float(page.PageSheet.CellsU("PageHeight").ResultIU)
        width = max(current_width, max_x + 0.75)
        height = max(current_height, max_y + 0.75)
        self._set_page_size(page, width, height)
