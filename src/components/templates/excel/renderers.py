from typing import Dict, Any, List
from .enums import (
    BorderStyle,
    NumberFormat,
    AlignmentType,
    SheetState,
    DataType,
    FilterType,
    SortOrder,
    ImagePosition
)
from .styles import ExcelStyles

class ExcelComponentsBase:

    @staticmethod
    def render_header(config: Dict[str, Any]) -> Dict[str, Any]:
        """
        Render Excel header row component.

        Configuration keys:
        - enabled (bool): Show/hide header row
        - text (str): Header text (if single-row header)
        - font_size (int): Font size (default 12)
        - colour (str): Text colour hex (default #FFFFFF)
        - text_colour (str): Alias for colour
        - bg_colour (str): Background colour hex (default #1F3A93)
        - cell_colour (str): Alias for bg_colour
        - bold (bool): Bold text (default True)
        - italic (bool): Italic text
        - underline (bool/str): Underline (True or Excel underline string)
        - alignment (str): left/center/right (default center)
        - font_name (str): Font name (default Calibri)
        - wrap_text (bool): Wrap text in header
        - height (float): Row height in points
        - border_style (str): Border style
        """
        if not config.get('enabled', True):
            return {'enabled': False}

        # Resolve colours
        text_color = config.get('text_color', config.get('color', '#FFFFFF'))
        bg_color = config.get('cell_color', config.get('bg_color', '#1F3A93'))

        return {
            'enabled': True,
            'text': config.get('text', None),
            'font_size': config.get('font_size', 12),
            'font_name': config.get('font_name', 'Calibri'),
            'bold': config.get('bold', True),
            'italic': config.get('italic', False),
            'underline': config.get('underline', False),
            'color': text_color,
            'bg_color': bg_color,
            'alignment': config.get('alignment', AlignmentType.CENTER.value),
            'wrap_text': config.get('wrap_text', False),
            'height': config.get('height', 20),  # points
            'border_style': config.get('border_style', BorderStyle.THIN.value)
        }

    @staticmethod
    def render_data_rows(config: Dict[str, Any]) -> Dict[str, Any]:
        """
        Render data row styling component.

        Config keys:
        - enabled (bool): Apply data styling
        - font_size (int): Font size (default 10)
        - font_name (str): Font name (default Calibri)
        - color (str): Text color hex
        - text_color (str): Alias for color
        - bg_color (str): Background color hex
        - cell_color (str): Alias for bg_color
        - bold (bool): Bold text
        - italic (bool): Italic text
        - underline (bool/str): Underline
        - alignment (str): Cell alignment
        - wrap_text (bool): Wrap text
        - height (float): Row height in points
        - number_format (str): Number format
        - alternating_colors (bool): Alternate row colors
        - alternate_color (str): Alternate row background color
        - border_style (str): Border style
        """
        if not config.get('enabled', True):
            return {'enabled': False}

        text_color = config.get('text_color', config.get('color', '#000000'))
        bg_color = config.get('cell_color', config.get('bg_color', '#FFFFFF'))

        return {
            'enabled': True,
            'font_size': config.get('font_size', 10),
            'font_name': config.get('font_name', 'Calibri'),
            'color': text_color,
            'bg_color': bg_color,
            'bold': config.get('bold', False),
            'italic': config.get('italic', False),
            'underline': config.get('underline', False),
            'alignment': config.get('alignment', AlignmentType.CENTER.value),
            'wrap_text': config.get('wrap_text', False),
            'height': config.get('height', 18),  # points
            'number_format': config.get('number_format', NumberFormat.GENERAL.value),
            'alternating_colors': config.get('alternating_colors', True),
            'alternate_color': config.get('alternate_color', '#F7FAFC'),
            'border_style': config.get('border_style', BorderStyle.THIN.value)
        }

    @staticmethod
    def render_totals_row(config: Dict[str, Any]) -> Dict[str, Any]:
        """
        Render totals row component (for sum/aggregate rows).

        Config keys:
        - enabled (bool): Show totals row
        - label (str): Label for first column (default "TOTAL")
        - font_size (int): Font size (default 11)
        - font_name (str): Font name
        - color (str): Text color hex
        - text_color (str): Alias for color
        - bg_color (str): Background color hex (default #E2E8F0)
        - cell_color (str): Alias for bg_color
        - bold (bool): Bold text (default True)
        - italic (bool): Italic text
        - underline (bool/str): Underline
        - alignment (str): Alignment
        - number_format (str): Number format for totals
        - border_style (str): Border style
        - formula_type (str): SUM/AVERAGE/COUNT/MIN/MAX
        """
        if not config.get('enabled', False):
            return {'enabled': False}

        text_color = config.get('text_color', config.get('color', '#000000'))
        bg_color = config.get('cell_color', config.get('bg_color', '#E2E8F0'))

        return {
            'enabled': True,
            'label': config.get('label', 'TOTAL'),
            'font_size': config.get('font_size', 11),
            'font_name': config.get('font_name', 'Calibri'),
            'color': text_color,
            'bg_color': bg_color,
            'bold': config.get('bold', True),
            'italic': config.get('italic', False),
            'underline': config.get('underline', False),
            'alignment': config.get('alignment', AlignmentType.LEFT.value),
            'number_format': config.get('number_format', NumberFormat.DECIMAL_2.value),
            'border_style': config.get('border_style', BorderStyle.THIN.value),
            'formula_type': config.get('formula_type', 'SUM')
        }

    @staticmethod
    def get_conditional_formatting(config: Dict[str, Any]) -> Dict[str, Any]:
        """
        Get conditional formatting rules.
        """
        if not config.get('enabled', False):
            return {'enabled': False}

        return {
            'enabled': True,
            'type': config.get('type', 'colorScale'),
            'color_min': config.get('color_min', '#F8CECC'),
            'color_mid': config.get('color_mid', '#FFFFFF'),
            'color_max': config.get('color_max', '#C3E6CB'),
            'formula': config.get('formula', None),
            'operator': config.get('operator', None),
            'value': config.get('value', None),
            'format': ExcelStyles.get_cell_format(config.get('format', {})),
            'range': config.get('range', None),
            'priority': config.get('priority', 1)
        }

    @staticmethod
    def get_data_validation(config: Dict[str, Any]) -> Dict[str, Any]:
        """
        Get data validation configuration.
        """
        if not config.get('enabled', False):
            return {'enabled': False}

        return {
            'enabled': True,
            'type': config.get('type', 'list'),
            'formula': config.get('formula', '"Option1,Option2,Option3"'),
            'operator': config.get('operator', None),
            'minimum': config.get('minimum', None),
            'maximum': config.get('maximum', None),
            'error_title': config.get('error_title', 'Invalid Entry'),
            'error_message': config.get('error_message', 'Please enter a valid value'),
            'error_style': config.get('error_style', 'stop'),
            'allow_blank': config.get('allow_blank', True),
            'show_dropdown': config.get('show_dropdown', True),
            'show_input': config.get('show_input', False),
            'input_title': config.get('input_title', None),
            'input_message': config.get('input_message', None),
            'range': config.get('range', None)
        }

    @staticmethod
    def get_freeze_panes_config(config: Dict[str, Any]) -> Dict[str, Any]:
        """
        Get freeze panes configuration.
        """
        return {
            'enabled': config.get('enabled', True),
            'freeze_rows': config.get('freeze_rows', 1),
            'freeze_cols': config.get('freeze_cols', 0),
            'cell': config.get('cell', None)
        }

    @staticmethod
    def get_sheet_protection(config: Dict[str, Any]) -> Dict[str, Any]:
        """
        Get sheet protection configuration.
        """
        if not config.get('enabled', False):
            return {'enabled': False}

        return {
            'enabled': True,
            'password': config.get('password', None),
            'allow_sort': config.get('allow_sort', True),
            'allow_filter': config.get('allow_filter', True),
            'allow_format_cells': config.get('allow_format_cells', False),
            'allow_insert_rows': config.get('allow_insert_rows', False),
            'allow_insert_cols': config.get('allow_insert_cols', False),
            'allow_delete_rows': config.get('allow_delete_rows', False),
            'allow_delete_cols': config.get('allow_delete_cols', False)
        }

    @staticmethod
    def render_sheet(config: Dict[str, Any]) -> Dict[str, Any]:
        """
        Sheet-level configuration (multi-sheet support).

        Config keys:
        - enabled (bool): Create/apply this sheet (default True)
        - name (str): Worksheet name/title
        - order (int): Sheet order (0-based). Lower comes first.
        - state (str): "visible" | "hidden" | "veryHidden"
        - hidden (bool): Convenience alias -> state="hidden"
        - very_hidden (bool): Convenience alias -> state="veryHidden"
        - tab_color (str): Optional hex
        """
        if not config.get("enabled", True):
            return {"enabled": False}

        state = config.get("state", SheetState.VISIBLE.value)
        if config.get("hidden", False):
            state = SheetState.HIDDEN.value

        if config.get("very_hidden", False):
            state = SheetState.VERY_HIDDEN.value

        return {
            "enabled": True,
            "name": config.get("name", "Sheet1"),
            "order": config.get("order", None),
            "state": state,
            "tab_color": config.get("tab_color", None),
        }

    @staticmethod
    def render_workbook_sheets(sheets: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
        """
        Normalize a list of sheets:
        - Calls render_sheet() for each entry
        - Filters disabled sheets
        - Sorts by 'order' if provided, otherwise preserves input order
        - Ensures non-empty names
        """
        normalized: List[Dict[str, Any]] = []
        for i, s in enumerate(sheets or []):
            rendered = ExcelComponentsBase.render_sheet(s or {})
            if not rendered.get("enabled", True):
                continue

            name = (rendered.get("name") or "").strip() or f"Sheet{i + 1}"
            rendered["name"] = name

            # Fallback ordering: if order not provided, keep the input order
            if rendered.get("order", None) is None:
                rendered["_implicit_order"] = i
            else:
                rendered["_implicit_order"] = rendered["order"]

            normalized.append(rendered)

        normalized.sort(key=lambda x: x["_implicit_order"])
        for s in normalized:
            s.pop("_implicit_order", None)

        return normalized

    @staticmethod
    def get_cell_data_type(config: Dict[str, Any]) -> Dict[str, Any]:
        """
        Define cell data type and validation.

        Config keys:
        - type (str): text/number/date/boolean (default "text")
        - allow_text (bool): Allow text input (default True)
        - allow_numbers (bool): Allow numeric input (default True)
        - allow_decimals (bool): Allow decimal numbers (default True)
        - allow_negative (bool): Allow negative numbers (default True)
        - date_format (str): Date format string (mm/dd/yyyy)
        - boolean_values (list): [true_value, false_value] (default [TRUE, FALSE])
        """
        data_type = config.get("type", DataType.TEXT.value)

        return {
            "type": data_type,
            "allow_text": config.get("allow_text", True),
            "allow_numbers": config.get("allow_numbers", True),
            "allow_decimals": config.get("allow_decimals", True),
            "allow_negative": config.get("allow_negative", True),
            "date_format": config.get("date_format", "mm/dd/yyyy"),
            "boolean_values": config.get("boolean_values", ["TRUE", "FALSE"]),
        }

    @staticmethod
    def get_cell_formula(config: Dict[str, Any]) -> Dict[str, Any]:
        """
        Define formulas

        Configuration keys:
        - enabled (bool): Enable formula (default False)
        - formula (str): Native Excel formula (e.g., "=SUM(A1:A10)", "=IF(B1>0,'Positive','Negative')")
        - description (str): Formula description for documentation
        - volatile (bool): Is formula volatile (recalculates always)
        - show_formula (bool): Display formula instead of the result (default False)

        Examples:
        - "=SUM(A1:A10)" - Sum range
        - "=AVERAGE(B:B)" - Average column
        - "=IF(C1>100,'High','Low')" - Conditional
        - "=Sheet2!A1+B1" - Cross-sheet reference
        - "=VLOOKUP(A1,Data,2,FALSE)" - Lookup with named range
        """
        if not config.get("enabled", False):
            return {"enabled": False}

        formula = config.get("formula", "")
        if not formula.startswith("="):
            formula = f"={formula}"

        return {
            "enabled": True,
            "formula": formula,
            "description": config.get("description", None),
            "volatile": config.get("volatile", False),
            "show_formula": config.get("show_formula", False),
        }

    @staticmethod
    def get_cross_sheet_reference(config: Dict[str, Any]) -> Dict[str, Any]:
        """
        Define cross-sheet cell references.

        Config keys:
        - enabled (bool): Enable cross-sheet reference (default False)
        - sheet_name (str): Name of target sheet
        - cell_ref (str): Cell reference (e.g., "A1", "B5:D10")
        - absolute (bool): Use absolute reference ($A$1) (default True)
        - reference_type (str): "single" (A1) or "range" (A1: D10)

        Examples:
        - sheet_name="Summary", cell_ref="A1" -> "=Summary!A1"
        - sheet_name="Data", cell_ref="A1:D10" -> "=Data!$A$1:$D$10"
        - Reference used in formulas: "=SUM(Sheet2!A1:A10)"
        """
        if not config.get("enabled", False):
            return {"enabled": False}

        sheet_name = config.get("sheet_name", "Sheet1")
        cell_ref = config.get("cell_ref", "A1")
        absolute = config.get("absolute", True)

        if absolute:
            if ":" in cell_ref:
                parts = cell_ref.split(":")
                cell_ref = f"${parts[0].strip()}:${parts[1].strip()}"
            else:
                cell_ref = f"${cell_ref.strip()}"

        reference = f"={sheet_name}!{cell_ref}"

        return {
            "enabled": True,
            "sheet_name": sheet_name,
            "cell_ref": cell_ref,
            "absolute": absolute,
            "reference": reference,
            "reference_type": config.get("reference_type", "single"),
        }

    @staticmethod
    def get_named_range(config: Dict[str, Any]) -> Dict[str, Any]:
        """
        Define named ranges for use in formulas.

        Config keys:
        - enabled (bool): Create the named range (default False)
        - name (str): Name of the range (e.g. "Sales_Data")
        - scope (str): "workbook" (visible everywhere) or "worksheet" (sheet-specific)
        - sheet_name (str): Sheet name (required if scope="worksheet")
        - range (str): Cell range (e.g. "A1:D100", "A:A" for entire column)
        - comment (str): Optional comment/description
        - hidden (bool): Hide (default False)

        Examples:
        - name="Sales_Data", range="A2:D100", scope="workbook"
        - name="Prices", range="Sheet2!B: B", scope="workbook"
        - name="Quarterly", range="C1:C4", scope="worksheet", sheet_name="Summary"
        """
        if not config.get("enabled", False):
            return {"enabled": False}

        name = config.get("name", "Range1")
        scope = config.get("scope", "workbook")
        range_ref = config.get("range", "A1")
        sheet_name = config.get("sheet_name", None)

        # Validate name
        if not name or not name[0].isalpha():
            raise ValueError(f"Named range name must start with letter: {name}")
        if " " in name:
            raise ValueError(f"Named range name cannot contain spaces: {name}")
        if len(name) > 255:
            raise ValueError(f"Named range name too long (max 255): {name}")

        if scope == "worksheet" and sheet_name:
            full_ref = f"{sheet_name}!{range_ref}"
        else:
            full_ref = range_ref

        return {
            "enabled": True,
            "name": name,
            "scope": scope,
            "sheet_name": sheet_name,
            "range": range_ref,
            "full_reference": full_ref,
            "comment": config.get("comment", None),
            "hidden": config.get("hidden", False),
        }

    @staticmethod
    def get_table_autofilter(config: dict) -> dict:
        """
        Enable AutoFilter on a table range.

        Configuration keys:
        - enabled (bool): Enable autofilter (default False)
        - range (str): Table range with headers (e.g., "A1:D100")
        - filter_type (str): "standard", "custom", or "advanced"
        - columns (list): Specific columns to filter (optional)
          Example: ["Name", "Sales", "Date"] or column indices [0, 1, 3]
        - default_sort (list): Default sort config
          [{
            "column": "Sales", # or column index
            "order": "desc", # "asc" or "desc"
            "sort_index": 1 # Priority (1=primary sort)
          }]
        - show_filter_buttons (bool): Show dropdown arrows (default True)
        """
        if not config.get("enabled", False):
            return {"enabled": False}

        range_ref = config.get("range", "A1:D1")
        filter_type = config.get("filter_type", FilterType.STANDARD.value)
        columns = config.get("columns", [])
        default_sort = config.get("default_sort", [])

        # Validate range format
        if ":" not in range_ref:
            raise ValueError(f"Invalid range format: {range_ref}. Use 'A1:D100'")

        return {
            "enabled": True,
            "range": range_ref,
            "filter_type": filter_type,
            "columns": columns,  # [] = all columns, ["A", "B"] = specific
            "default_sort": sorted(default_sort, key=lambda x: x.get("sort_index", 999)),
            "show_filter_buttons": config.get("show_filter_buttons", True),
            "description": "AutoFilter with dropdown arrows in header row"
        }

    @staticmethod
    def get_table_sorting(config: dict) -> dict:
        """
        Configure default table sorting (applied when table created).

        Configuration keys:
        - enabled (bool): Enable sorting (default False)
        - columns (list): Sort columns in priority order
          [{
            "column": "Sales", # Column name or index
            "order": "desc", # "asc" or "desc"
            "sort_index": 1, # 1=primary, 2=secondary, etc.
            "case_sensitive": False # Text sorting case sensitivity
          }]
        - default_sort_state (str): "none", "asc", or "desc"
        """
        if not config.get("enabled", False):
            return {"enabled": False}

        columns = config.get("columns", [])

        # Validate sort config
        for col_config in columns:
            if "column" not in col_config:
                raise ValueError("Sort column config missing 'column' key")
            if col_config.get("order") not in [SortOrder.ASCENDING.value, SortOrder.DESCENDING.value]:
                raise ValueError(f"Invalid sort order: {col_config.get('order')}")

        # Sort by priority
        sorted_columns = sorted(columns, key=lambda x: x.get("sort_index", 999))

        return {
            "enabled": True,
            "columns": sorted_columns,
            "total_sort_levels": len(sorted_columns),
            "description": f"Multi-level sort: {len(sorted_columns)} columns"
        }

    @staticmethod
    def get_freeze_panes(config: dict) -> dict:
        """
        Freeze rows and/or columns to keep headers visible during scrolling.

        Config keys:
        - enabled (bool): Enable freeze panes (default False)
        - freeze_type (str): "rows", "columns", "both", or "none"
        - freeze_rows (int): Number of rows to freeze from top (default 1)
        - freeze_columns (int): Number of columns to freeze from left (default 0)
        - freeze_cell (str): Alternative: specify cell (e.g., "B2" freezes above and left)

        Common patterns:
        - freeze_rows=1 (header row)
        - freeze_columns=1 (first column with IDs)
        - freeze_cell="B2" (freeze row 1 and column A)
        - freeze_rows=2, freeze_columns=1 (freeze 2 rows + 1 column)
        """
        if not config.get("enabled", False):
            return {"enabled": False}

        freeze_type = config.get("freeze_type", "both").lower()
        freeze_cell = config.get("freeze_cell")

        # If freeze_cell specified, parse it
        if freeze_cell:
            # Extract column letter and row number from the cell (e.g., "B2" -> col=2, row=2)
            import re
            match = re.match(r'([A-Z]+)(\d+)', freeze_cell.upper())
            if match:
                col_letter = match.group(1)
                row_num = int(match.group(2))
                # Convert column letter to number (A=1, B=2, etc)
                freeze_columns = sum(
                    (ord(char) - ord('A') + 1) * 26 ** i
                    for i, char in enumerate(reversed(col_letter))
                )
                freeze_rows = row_num - 1  # 0-indexed
            else:
                raise ValueError(f"Invalid cell reference: {freeze_cell}")
        else:
            freeze_rows = config.get("freeze_rows", 1)
            freeze_columns = config.get("freeze_columns", 0)

        # Validate
        if freeze_rows < 0 or freeze_columns < 0:
            raise ValueError("freeze_rows and freeze_columns must be >= 0")

        # Map freeze_type to actual panes
        if freeze_type == "rows":
            freeze_columns = 0
        elif freeze_type == "columns":
            freeze_rows = 0
        elif freeze_type == "both":
            pass  # Use both values
        elif freeze_type == "none":
            freeze_rows = 0
            freeze_columns = 0

        return {
            "enabled": True,
            "freeze_type": freeze_type,
            "freeze_rows": freeze_rows,
            "freeze_columns": freeze_columns,
            "freeze_cell": freeze_cell or f"{chr(65 + freeze_columns - 1) if freeze_columns else 'A'}{freeze_rows + 1}",
            "description": f"Freeze {freeze_rows} rows, {freeze_columns} columns"
        }

    @staticmethod
    def get_header_image(config: dict) -> dict:
        """
        Add image/logo to the header area

        Configuration keys:
        - enabled (bool): Add image (default False)
        - image_path (str): Path to image file (PNG, JPG, GIF, etc)
            OR image_url (str): URL to remote image
        - position (str): "topleft", "topcenter", "topright", "center", etc
        - row (int): Starting row (default 1)
        - column (int): Starting column (default 1 = A)
        - width (float): Image width in cm (default 3.0)
        - height (float): Image height in cm (default 1.5)
        - maintain_aspect_ratio (bool): Keep aspect ratio (default True)
        - anchor_type (str): "cell" (move/resize with cell) or "absolute" (fixed)
        - transparency (float): 0-100 (0=opaque, 100=transparent)
        - rotation (int): Rotation in degrees (0-360)
        - description (str): Alt text for accessibility
        """
        if not config.get("enabled", False):
            return {"enabled": False}

        image_path = config.get("image_path")
        image_url = config.get("image_url")

        if not image_path and not image_url:
            raise ValueError("Must provide either image_path or image_url")

        position = config.get("position", ImagePosition.TOP_LEFT.value)
        row = config.get("row", 1)
        column = config.get("column", 1)
        width = config.get("width", 3.0)  # cm
        height = config.get("height", 1.5)  # cm
        maintain_aspect = config.get("maintain_aspect_ratio", True)
        anchor_type = config.get("anchor_type", "cell")
        transparency = config.get("transparency", 0)
        rotation = config.get("rotation", 0)

        # Validate inputs
        if not 0 <= transparency <= 100:
            raise ValueError("Transparency must be 0-100")
        if not 0 <= rotation <= 360:
            raise ValueError("Rotation must be 0-360")
        if width <= 0 or height <= 0:
            raise ValueError("Width and height must be positive")

        return {
            "enabled": True,
            "image_source": image_path or image_url,
            "source_type": "file" if image_path else "url",
            "position": position,
            "row": row,
            "column": column,
            "width_cm": width,
            "height_cm": height,
            "maintain_aspect_ratio": maintain_aspect,
            "anchor_type": anchor_type,
            "transparency_percent": transparency,
            "rotation_degrees": rotation,
            "description": config.get("description", "Header image/logo"),
            "cell_reference": f"{chr(64 + column)}{row}"  # A1, B2, etc
        }