from enum import Enum

class BorderStyle(Enum):
    NONE = "none"
    THIN = "thin"
    MEDIUM = "medium"
    THICK = "thick"
    DOTTED = "dotted"
    DASHED = "dashed"
    DOUBLE = "double"

class NumberFormat(Enum):
    GENERAL = "General"
    INTEGER = "0"
    DECIMAL_2 = "#,##0.00"
    DECIMAL_4 = "#,##0.0000"
    PERCENTAGE = "0.00%"
    CURRENCY_USD = "$#,##0.00"
    CURRENCY_EUR = "€ #,##0.00"
    DATE_SHORT = "mm/dd/yyyy"
    DATE_LONG = "mmmm d, yyyy"
    TIME_24 = "hh:mm:ss"
    TIME_12 = "h:mm AM/PM"
    DATETIME = "mm/dd/yyyy hh:mm"

class AlignmentType(Enum):
    LEFT = "left"
    CENTER = "center"
    RIGHT = "right"
    JUSTIFY = "justify"
    DISTRIBUTED = "distributed"

class SheetState(Enum):
    VISIBLE = "visible"
    HIDDEN = "hidden"
    VERY_HIDDEN = "veryHidden"

class DataType(Enum):
    TEXT = "text"
    NUMBER = "number"
    DATE = "date"
    BOOLEAN = "boolean"

class FilterType(Enum):
    STANDARD = "standard"
    CUSTOM = "custom"
    ADVANCED = "advanced"

class SortOrder(Enum):
    ASCENDING = "asc"
    DESCENDING = "desc"

class ImagePosition(Enum):
    TOP_LEFT = "topleft"
    TOP_CENTER = "topcenter"
    TOP_RIGHT = "topright"
    CENTER = "center"
    BOTTOM_LEFT = "bottomleft"
    BOTTOM_CENTER = "bottomcenter"
    BOTTOM_RIGHT = "bottomright"