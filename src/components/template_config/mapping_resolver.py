from typing import Any, Dict, List, Optional
from jsonpath_ng import parse as jsonpath_parse


class PlaceholderMapping:

    def __init__(
            self,
            placeholder: str,
            data_path: str,
            transform: Optional[str] = None,
            default_value: Any = None
    ):
        self.placeholder = placeholder
        self.data_path = data_path
        self.transform = transform
        self.default_value = default_value


class MappingResolver:

    TRANSFORMS = {
        "uppercase": lambda x: str(x).upper(),
        "lowercase": lambda x: str(x).lower(),
        "titlecase": lambda x: str(x).title(),
        "capitalize": lambda x: str(x).capitalize(),
    }

    @staticmethod
    def register_transform(name: str, func):
        MappingResolver.TRANSFORMS[name] = func

    @staticmethod
    def resolve_path(data: Dict[str, Any], path: str) -> Any:
        """
        Resolve JSONPath expression in data dict.
        Supports:
            - "customer.name" → data["customer"]["name"]
            - "customers[0].name" → data["customers"][0]["name"]
            - "items[*].price" → [item["price"] for item in data["items"]]
        """
        try:
            expr = jsonpath_parse(path)
            matches = [match.value for match in expr.find(data)]
            if not matches:
                return None
            return matches[0] if len(matches) == 1 else matches
        except Exception:
            return None

    @staticmethod
    def apply_transform(value: Any, transform: Optional[str]) -> Any:
        if not transform or value is None:
            return value

        parts = transform.split(":")
        transform_name = parts[0].lower().strip()

        if transform_name in MappingResolver.TRANSFORMS:
            return MappingResolver.TRANSFORMS[transform_name](value)

        # Date format: "date:DD.MM.YYYY"
        if transform_name == "date":
            from datetime import datetime
            if len(parts) < 2:
                return value
            fmt = ":".join(parts[1:])
            if isinstance(value, str):
                # Try common formats
                for parse_fmt in ["%Y-%m-%d", "%d.%m.%Y", "%d/%m/%Y", "%Y/%m/%d"]:
                    try:
                        dt = datetime.strptime(value, parse_fmt)
                        return dt.strftime(fmt)
                    except ValueError:
                        continue
            return value

        if transform_name == "currency":
            symbol = parts[1] if len(parts) > 1 else "€"
            try:
                num = float(value)
                return f"{symbol} {num:,.2f}"
            except (ValueError, TypeError):
                return value

        return value

    @staticmethod
    def resolve_all(
            data: Dict[str, Any],
            mappings: List[PlaceholderMapping],
            auto_fallback: bool = True
    ) -> Dict[str, Any]:
        resolved = {}

        for mapping in mappings:
            value = MappingResolver.resolve_path(data, mapping.data_path)
            if value is None and auto_fallback:
                value = MappingResolver.resolve_path(data, mapping.placeholder)

            if value is None:
                value = mapping.default_value

            value = MappingResolver.apply_transform(value, mapping.transform)

            resolved[mapping.placeholder] = value

        return resolved

    @staticmethod
    def build_table_context(
            data: Dict[str, Any],
            table_mappings: List[Dict[str, Any]],
            array_path: str
    ) -> List[Dict[str, Any]]:
        """
        Convert array data to table rows.
        """
        array_data = MappingResolver.resolve_path(data, array_path)
        if not isinstance(array_data, list):
            return []

        rows = []
        for item in array_data:
            row = {}
            for col_map in table_mappings:
                col_name = col_map["column"]
                item_path = col_map.get("path", col_map["column"])
                value = MappingResolver.resolve_path(item, item_path)
                value = MappingResolver.apply_transform(value, col_map.get("transform"))
                row[col_name] = value or ""
            rows.append(row)

        return rows

    @staticmethod
    def build_context_from_dict(
            data: Dict[str, Any],
            mapping_dict: Optional[List[Dict[str, Any]]] = None,
            table_mappings: Optional[List[Dict[str, Any]]] = None,
            table_placeholder: Optional[str] = None,
            auto_fallback: bool = True
    ) -> Dict[str, Any]:
        """
        Build raw dict format (e.g. from JSON input).
        """
        mappings = [
            PlaceholderMapping(
                placeholder=m["placeholder"],
                data_path=m.get("path", m["placeholder"]),
                transform=m.get("transform"),
                default_value=m.get("default")
            )
            for m in mapping_dict or []
        ]
        context = MappingResolver.resolve_all(data, mappings, auto_fallback)

        if table_mappings and table_placeholder:
            table_array_path = table_mappings[0].get("array_path", "table_data[*]")
            table_rows = MappingResolver.build_table_context(data, table_mappings, table_array_path)
            context[table_placeholder] = table_rows

        return context