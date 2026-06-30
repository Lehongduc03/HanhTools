# -*- coding: utf-8 -*-
__title__ = "Insulation Rule Editor"
__author__ = "TD"
__version__ = "V4.4"
__persistentengine__ = True
__doc__ = """
Rule Editor modeless cho pyRevit IronPython 2.7.

GHI CHÚ TIẾNG VIỆT:
- UI mở bằng window.Show(), không khóa Revit.
- Các thao tác đọc/ghi Revit API dùng ExternalEvent.
- Người dùng khai báo cột rule động trong bảng Parameter Column Mapping.
- Rule DataGrid tự sinh cột theo mapping.
- Nếu ô dynamic trong rule trống thì bỏ qua điều kiện đó.
- Nếu mapping disabled hoặc thiếu RevitFieldName/Operator thì bỏ qua cột đó.
- Save/load gồm 2 file CSV:
    1) insulation_rule_mapping.csv
    2) insulation_rules.csv
- V4.3: giữ nguyên UI hiện tại, loại bỏ cột MinDN / MaxDN khỏi Rule DataGrid và file rules mới.
  Nếu cần lọc size, hãy dùng Parameter Column Mapping với RevitFieldName = Diameter / Width / Height / Overall Size.
  InsulationTypeName dropdown được lọc theo ElementType của từng dòng rule:
  Pipe/Pipe Fitting/Pipe Accessory/Flex Pipe -> Pipe Insulation Type.
  Duct/Duct Fitting/Duct Accessory/Flex Duct -> Duct Insulation Type.
  Không thay đổi cấu trúc so sánh dynamic mapping.
- V4.4: bổ sung cơ chế sửa chữa insulation hiện có:
  Nếu element đã có insulation và đúng type + thickness theo rule thì skip.
  Nếu element đã có insulation nhưng sai type hoặc sai thickness thì tool xóa insulation cũ và tạo lại theo rule trong SubTransaction.
  UI giữ nguyên, workflow giữ nguyên.
"""

import clr
import os
import csv
import re
import traceback
from StringIO import StringIO

clr.AddReference("RevitAPI")
clr.AddReference("RevitAPIUI")
clr.AddReference("PresentationFramework")
clr.AddReference("PresentationCore")
clr.AddReference("WindowsBase")

from System import Action
from System.Text import Encoding
from System.IO import File
from System.Collections.ObjectModel import ObservableCollection

from System.Windows import Setter, DataTemplate, FrameworkElementFactory
from System.Windows.Controls import (
    ComboBox,
    DataGridCheckBoxColumn,
    DataGridComboBoxColumn,
    DataGridTextColumn,
    DataGridLength
)
from System.Windows.Data import Binding as WpfBinding, BindingMode, UpdateSourceTrigger
from System.Windows.Threading import DispatcherPriority

from Microsoft.Win32 import OpenFileDialog, SaveFileDialog

from Autodesk.Revit.DB import *
from Autodesk.Revit.UI import TaskDialog, IExternalEventHandler, ExternalEvent

try:
    from Autodesk.Revit.DB.Plumbing import PipeInsulation, PipeInsulationType
except:
    PipeInsulation = None
    PipeInsulationType = None

try:
    from Autodesk.Revit.DB.Mechanical import DuctInsulation, DuctInsulationType
except:
    DuctInsulation = None
    DuctInsulationType = None

from pyrevit import forms, script


# =============================================================================
# REVIT CONTEXT
# =============================================================================

uidoc = __revit__.ActiveUIDocument
doc = uidoc.Document if uidoc else None
app = __revit__.Application
output = script.get_output()


# =============================================================================
# PATHS
# =============================================================================

SCRIPT_DIR = os.path.dirname(__file__)
XAML_FILE = os.path.join(SCRIPT_DIR, "RuleEditor.xaml")
DEFAULT_MAPPING_CSV = os.path.join(SCRIPT_DIR, "insulation_rule_mapping.csv")
DEFAULT_RULE_CSV = os.path.join(SCRIPT_DIR, "insulation_rules.csv")


# =============================================================================
# PERSISTENT GLOBALS
# =============================================================================

_RULE_EDITOR_WINDOW = None
_RULE_EDITOR_HANDLER = None
_RULE_EDITOR_EVENT = None


# =============================================================================
# CSV HEADERS
# =============================================================================

MAPPING_FIELDS = [
    "Enabled",
    "RuleColumnName",
    "RevitFieldName",
    "Operator",
    "ValueSource",
    "ValueType",
    "Note"
]

FIXED_RULE_FIELDS = [
    "Enabled",
    "RuleId",
    "ElementType",
    "MinDN",
    "MaxDN",
    "ThicknessMM",
    "InsulationTypeName",
    "Priority",
    "Note"
]

FIXED_RULE_FIELD_SET = {}
for _f in FIXED_RULE_FIELDS:
    FIXED_RULE_FIELD_SET[_f.lower()] = True


# =============================================================================
# OPTIONS
# =============================================================================

ELEMENT_TYPE_OPTIONS = [
    "Pipe",
    "Pipe Fitting",
    "Pipe Accessory",
    "Flex Pipe",
    "Duct",
    "Duct Fitting",
    "Duct Accessory",
    "Flex Duct",
    "All Pipe",
    "All Duct",
    "Both",
    "All"
]

OPERATOR_OPTIONS = [
    "",
    "Equals",
    "NotEquals",
    "Contains",
    "NotContains",
    "StartsWith",
    "EndsWith",
    "GreaterThan",
    "GreaterOrEqual",
    "LessThan",
    "LessOrEqual",
    "IsEmpty",
    "IsNotEmpty"
]

VALUE_SOURCE_OPTIONS = [
    "RevitDropdown",
    "Manual",
    "ManualOrDropdown"
]

VALUE_TYPE_OPTIONS = [
    "Auto",
    "Text",
    "Number",
    "LengthMM",
    "SizeText"
]

BASE_REVIT_FIELD_OPTIONS = [
    "",
    "System Name",
    "System Abbreviation",
    "System Classification",
    "Level Name",
    "Workset",
    "Phase",
    "View Name",
    "Category",
    "Family Name",
    "Type Name",
    "InsulationTypeName",
    "Diameter",
    "Width",
    "Height",
    "Length",
    "Overall Size",
    "Room Name",
    "Space Name",
    "Installation Area",
    "Service Type",
    "Insulation Area",
    "SI_InstallationArea",
    "TGA_SystemCode",
    "Comments",
    "Mark",
    "Type Comments"
]

NUMERIC_OPERATORS = [
    "GreaterThan",
    "GreaterOrEqual",
    "LessThan",
    "LessOrEqual"
]

NUMERIC_FIELDS = [
    "Diameter",
    "Width",
    "Height",
    "Length",
    "Overall Size",
    "ThicknessMM",
    "Priority"
]

DEFAULT_TRADE_VALUES = [
    "",
    "HZG",
    "SAN",
    "RLT",
    "SPR",
    "CHW",
    "CWS",
    "CWR"
]

# Nếu TRUE: mapping disabled vẫn hiện cột trong Rule DataGrid, nhưng không dùng để match.
SHOW_DISABLED_MAPPING_COLUMNS = True

# V4.4:
# Nếu TRUE: element đã có insulation sai rule sẽ được sửa bằng cách xóa insulation cũ và tạo lại theo rule.
# Nếu insulation hiện có đã đúng type + thickness thì tool skip để tránh sửa dư.
REPAIR_EXISTING_INSULATION_BY_RULE = True

# Giữ biến cũ để tương thích, nhưng V4.4 không còn dùng kiểu skip toàn bộ existing insulation.
REPLACE_EXISTING_INSULATION = False

# Sai số cho so sánh thickness insulation, đơn vị mm.
INSULATION_THICKNESS_TOLERANCE_MM = 0.5

# -----------------------------------------------------------------------------
# TỐI ƯU KHI MỞ TOOL
# -----------------------------------------------------------------------------
# Giữ nguyên cấu trúc UI của V3.5, chỉ tối ưu cách load để form lên nhanh hơn.
# False = khi mở UI chỉ load XAML + CSV, không tự quét Revit ngay lúc startup.
# Người dùng bấm Refresh Elements hoặc Refresh Dropdown Values khi cần lấy dữ liệu mới.
AUTO_REFRESH_DROPDOWNS_ON_STARTUP = False

# Giới hạn số element dùng để tạo dropdown khi bấm Refresh Dropdown Values.
# Đặt None nếu muốn quét hết, nhưng với model lớn sẽ chậm.
MAX_DROPDOWN_ELEMENT_SCAN = 500

# False = không collect toàn bộ View trong project khi tạo dropdown View Name.
# Tool vẫn thêm Active View hiện tại vào dropdown. Project lớn collect all views rất tốn thời gian.
INCLUDE_ALL_VIEWS_IN_DROPDOWN = False

# Key nội bộ để lưu dropdown theo ElementType.
# Không ghi ra CSV, chỉ dùng cho UI Rule DataGrid.
DROPDOWN_BY_ELEMENT_TYPE_KEY = "__ByElementType__"


# =============================================================================
# BASIC HELPERS
# =============================================================================

def alert(message):
    TaskDialog.Show(__title__, message)


def to_unicode(value):
    if value is None:
        return u""

    try:
        if isinstance(value, unicode):
            return value
    except:
        pass

    try:
        return unicode(value, "utf-8")
    except:
        pass

    try:
        return unicode(value)
    except:
        return u""


def fix_mojibake(text):
    """
    Sua loi ma hoa hay gap khi CSV bi doc sai encoding.
    Vi du: mmÃ¸ -> mmø. Neu khong sua duoc thi giu nguyen.
    """
    try:
        if u"Ã" in text or u"Â" in text:
            return text.encode("latin-1").decode("utf-8")
    except:
        pass
    return text


def clean_text(value):
    return fix_mojibake(to_unicode(value)).strip()


def bool_from_csv(value):
    text = clean_text(value).lower()
    return text in ["true", "1", "yes", "y", "x"]


def bool_to_csv(value):
    return "TRUE" if bool(value) else "FALSE"


def try_float(value):
    """
    Chuyen text sang so.
    Ho tro gia tri co suffix Revit/CSV nhu:
    - 100 mm
    - 100 mmø-80 mmø
    - DN50
    - 12,5 mm
    Ham lay so dau tien tim duoc neu float truc tiep that bai.
    """
    text = clean_text(value).replace(",", ".")
    if text == "":
        return None

    try:
        return float(text)
    except:
        pass

    try:
        m = re.search(r"[-+]?\d+(?:\.\d+)?", text)
        if m:
            return float(m.group(0))
    except:
        pass

    return None


def strings_equal(a, b):
    return clean_text(a).lower() == clean_text(b).lower()


def contains_text(a, b):
    return clean_text(b).lower() in clean_text(a).lower()


def starts_with_text(a, b):
    return clean_text(a).lower().startswith(clean_text(b).lower())


def ends_with_text(a, b):
    return clean_text(a).lower().endswith(clean_text(b).lower())


def element_id_value(element_id):
    try:
        return element_id.IntegerValue
    except:
        pass

    try:
        return element_id.Value
    except:
        return None


def get_element_name(elem):
    if elem is None:
        return ""

    try:
        name = elem.Name
        if name:
            return clean_text(name)
    except:
        pass

    try:
        name = Element.Name.GetValue(elem)
        if name:
            return clean_text(name)
    except:
        pass

    for bip_name in ["SYMBOL_NAME_PARAM", "ALL_MODEL_TYPE_NAME"]:
        try:
            bip = getattr(BuiltInParameter, bip_name)
            p = elem.get_Parameter(bip)
            if p and p.HasValue:
                value = p.AsString()
                if value:
                    return clean_text(value)

                value = p.AsValueString()
                if value:
                    return clean_text(value)
        except:
            pass

    return ""


def length_internal_to_mm(value):
    try:
        return float(value) * 304.8
    except:
        return 0.0


def mm_to_internal(value_mm):
    try:
        return float(value_mm) / 304.8
    except:
        return 0.0


def format_number(value):
    if value is None:
        return ""

    try:
        value = float(value)
        if abs(value - round(value)) < 0.000001:
            return str(int(round(value)))
        return "{0:.2f}".format(value)
    except:
        return clean_text(value)


def safe_binding_name(column_name):
    """
    Chuyển tên cột CSV thành tên property an toàn để WPF Binding.
    Ví dụ:
        System Name -> Dyn_System_Name
        Level.Name  -> Dyn_Level_Name
    """
    text = clean_text(column_name)
    text = re.sub("[^0-9a-zA-Z_]", "_", text)

    if text == "":
        text = "Column"

    if text[0].isdigit():
        text = "C_" + text

    return "Dyn_" + text


def add_unique_text(data_dict, value):
    text = clean_text(value)
    if text:
        data_dict[text] = True


def sorted_keys(data_dict):
    keys = list(data_dict.keys())
    keys.sort()
    return keys


def reset_observable(collection, values):
    collection.Clear()
    for value in values:
        collection.Add(value)


def yes_no(message):
    try:
        return forms.alert(message, yes=True, no=True)
    except:
        return False


def dg_width(value):
    """
    WPF DataGridColumn.Width cần DataGridLength.
    Không dùng trực tiếp số int, nếu không IronPython báo: expected DataGridLength, got int.
    """
    try:
        return DataGridLength(float(value))
    except:
        return DataGridLength(120.0)


# =============================================================================
# CSV
# =============================================================================

def read_csv_dicts(path):
    rows = []

    if not path or not os.path.exists(path):
        return rows

    text = File.ReadAllText(path, Encoding.UTF8)

    if text.startswith(u"\ufeff"):
        text = text[1:]

    buffer = StringIO(text.encode("utf-8"))
    reader = csv.DictReader(buffer)

    for row in reader:
        clean_row = {}
        for key, value in row.items():
            clean_row[clean_text(key)] = clean_text(value)
        rows.append(clean_row)

    return rows


def read_csv_headers(path):
    if not path or not os.path.exists(path):
        return []

    text = File.ReadAllText(path, Encoding.UTF8)

    if text.startswith(u"\ufeff"):
        text = text[1:]

    buffer = StringIO(text.encode("utf-8"))
    reader = csv.reader(buffer)

    try:
        headers = reader.next()
    except:
        return []

    result = []
    for h in headers:
        result.append(clean_text(h))

    return result


def csv_escape(value):
    text = clean_text(value)
    text = text.replace('"', '""')

    if "," in text or '"' in text or "\n" in text or "\r" in text:
        return '"' + text + '"'

    return text


def write_csv_dicts(path, headers, rows):
    lines = []
    lines.append(",".join(headers))

    for row in rows:
        values = []
        for header in headers:
            values.append(csv_escape(row.get(header, "")))
        lines.append(",".join(values))

    File.WriteAllText(path, u"\n".join(lines), Encoding.UTF8)


def pick_csv_file(title):
    dlg = OpenFileDialog()
    dlg.Title = title
    dlg.Filter = "CSV Files (*.csv)|*.csv|All Files (*.*)|*.*"
    dlg.Multiselect = False

    if dlg.ShowDialog():
        return dlg.FileName

    return None


def save_csv_file(title, default_name):
    dlg = SaveFileDialog()
    dlg.Title = title
    dlg.Filter = "CSV Files (*.csv)|*.csv|All Files (*.*)|*.*"
    dlg.FileName = default_name

    if dlg.ShowDialog():
        return dlg.FileName

    return None


# =============================================================================
# DATA CLASSES
# =============================================================================

class MappingItem(object):
    """
    Một dòng mapping đại diện cho một cột động trong Rule DataGrid.
    """
    def __init__(self):
        self.Enabled = True
        self.RuleColumnName = ""
        self.RevitFieldName = ""
        self.Operator = "Equals"
        self.ValueSource = "RevitDropdown"
        self.ValueType = "Auto"
        self.Note = ""
        self.BindingName = ""

    def update_binding_name(self):
        self.BindingName = safe_binding_name(self.RuleColumnName)

    def to_dict(self):
        return {
            "Enabled": bool_to_csv(self.Enabled),
            "RuleColumnName": clean_text(self.RuleColumnName),
            "RevitFieldName": clean_text(self.RevitFieldName),
            "Operator": clean_text(self.Operator),
            "ValueSource": clean_text(self.ValueSource),
            "ValueType": clean_text(self.ValueType),
            "Note": clean_text(self.Note)
        }


class RuleItem(object):
    """
    Một dòng rule.
    Các cột động được lưu bằng setattr theo BindingName từ MappingItem.
    """
    def __init__(self):
        self.Enabled = True
        self.RuleId = ""
        self.ElementType = "Both"
        # V4.3: MinDN/MaxDN được giữ nội bộ để đọc file CSV cũ, nhưng không còn hiển thị/lưu mới.
        self.MinDN = ""
        self.MaxDN = ""
        self.ThicknessMM = ""
        self.InsulationTypeName = ""
        self.Priority = "100"
        self.Note = ""

    def set_dynamic_value(self, mapping, value):
        mapping.update_binding_name()
        setattr(self, mapping.BindingName, clean_text(value))

    def get_dynamic_value(self, mapping):
        mapping.update_binding_name()
        try:
            return clean_text(getattr(self, mapping.BindingName))
        except:
            return ""

    def to_dict(self, mappings):
        row = {
            "Enabled": bool_to_csv(self.Enabled),
            "RuleId": clean_text(self.RuleId),
            "ElementType": clean_text(self.ElementType),
            "ThicknessMM": clean_text(self.ThicknessMM),
            "InsulationTypeName": clean_text(self.InsulationTypeName),
            "Priority": clean_text(self.Priority),
            "Note": clean_text(self.Note)
        }

        for mapping in mappings:
            col_name = clean_text(mapping.RuleColumnName)
            if col_name:
                row[col_name] = self.get_dynamic_value(mapping)

        return row


class ElementPreviewItem(object):
    """
    Dòng preview trong tab Active View Elements.
    Custom field cũng được gắn bằng setattr.
    """
    def __init__(self):
        self.ElementId = ""
        self.Category = ""
        self.SystemName = ""
        self.SystemAbbreviation = ""
        self.SystemClassification = ""
        self.LevelName = ""
        self.Workset = ""
        self.Phase = ""
        self.FamilyName = ""
        self.TypeName = ""
        self.Diameter = ""
        self.Width = ""
        self.Height = ""
        self.Length = ""
        self.OverallSize = ""
        self.InsulationExisting = ""
        self.Comments = ""
        self.Mark = ""

    def search_blob(self):
        values = []
        for key, value in self.__dict__.items():
            values.append(clean_text(value))
        return " ".join(values).lower()


def mapping_from_dict(row):
    item = MappingItem()
    item.Enabled = bool_from_csv(row.get("Enabled", "TRUE"))
    item.RuleColumnName = clean_text(row.get("RuleColumnName", ""))
    item.RevitFieldName = clean_text(row.get("RevitFieldName", ""))
    item.Operator = clean_text(row.get("Operator", "Equals"))
    item.ValueSource = clean_text(row.get("ValueSource", "RevitDropdown"))
    item.ValueType = clean_text(row.get("ValueType", "Auto"))
    item.Note = clean_text(row.get("Note", ""))

    if item.Operator not in OPERATOR_OPTIONS:
        item.Operator = "Equals"

    if item.ValueSource not in VALUE_SOURCE_OPTIONS:
        item.ValueSource = "RevitDropdown"

    if item.ValueType not in VALUE_TYPE_OPTIONS:
        item.ValueType = "Auto"

    item.update_binding_name()
    return item


def rule_from_dict(row, mappings):
    item = RuleItem()
    item.Enabled = bool_from_csv(row.get("Enabled", "TRUE"))
    item.RuleId = clean_text(row.get("RuleId", ""))
    item.ElementType = clean_text(row.get("ElementType", "Both"))
    item.MinDN = clean_text(row.get("MinDN", ""))
    item.MaxDN = clean_text(row.get("MaxDN", ""))
    item.ThicknessMM = clean_text(row.get("ThicknessMM", ""))
    item.InsulationTypeName = clean_text(row.get("InsulationTypeName", ""))
    item.Priority = clean_text(row.get("Priority", "100"))
    item.Note = clean_text(row.get("Note", ""))

    if item.ElementType not in ELEMENT_TYPE_OPTIONS:
        item.ElementType = "Both"

    for mapping in mappings:
        col_name = clean_text(mapping.RuleColumnName)
        if col_name:
            item.set_dynamic_value(mapping, row.get(col_name, ""))

    return item


# =============================================================================
# REVIT VALUE READERS
# =============================================================================

def get_param_as_text(elem, param_name):
    if elem is None or not param_name:
        return ""

    try:
        p = elem.LookupParameter(param_name)
    except:
        p = None

    if p is None:
        return ""

    try:
        if not p.HasValue:
            return ""
    except:
        pass

    try:
        if p.StorageType == StorageType.String:
            return clean_text(p.AsString())
    except:
        pass

    try:
        if p.StorageType == StorageType.Double:
            val = p.AsValueString()
            if val:
                return clean_text(val)
            return clean_text(p.AsDouble())
    except:
        pass

    try:
        if p.StorageType == StorageType.Integer:
            val = p.AsValueString()
            if val:
                return clean_text(val)
            return clean_text(p.AsInteger())
    except:
        pass

    try:
        if p.StorageType == StorageType.ElementId:
            ref_elem = elem.Document.GetElement(p.AsElementId())
            if ref_elem:
                return get_element_name(ref_elem)
            return clean_text(element_id_value(p.AsElementId()))
    except:
        pass

    try:
        val = p.AsValueString()
        if val:
            return clean_text(val)
    except:
        pass

    return ""


def get_bip_param_text(elem, bip_names):
    if elem is None:
        return ""

    for bip_name in bip_names:
        try:
            bip = getattr(BuiltInParameter, bip_name)
            p = elem.get_Parameter(bip)
            if p and p.HasValue:
                val = p.AsString()
                if val:
                    return clean_text(val)

                val = p.AsValueString()
                if val:
                    return clean_text(val)

                if p.StorageType == StorageType.Double:
                    return clean_text(p.AsDouble())

                if p.StorageType == StorageType.Integer:
                    return clean_text(p.AsInteger())

                if p.StorageType == StorageType.ElementId:
                    ref_elem = elem.Document.GetElement(p.AsElementId())
                    if ref_elem:
                        return get_element_name(ref_elem)
                    return clean_text(element_id_value(p.AsElementId()))
        except:
            pass

    return ""


def get_length_param_mm(elem, bip_names, param_names):
    for bip_name in bip_names:
        try:
            bip = getattr(BuiltInParameter, bip_name)
            p = elem.get_Parameter(bip)
            if p and p.HasValue:
                return length_internal_to_mm(p.AsDouble())
        except:
            pass

    for name in param_names:
        try:
            p = elem.LookupParameter(name)
            if p and p.HasValue:
                return length_internal_to_mm(p.AsDouble())
        except:
            pass

    return None


def get_category_name(elem):
    try:
        if elem.Category:
            return clean_text(elem.Category.Name)
    except:
        pass
    return ""


def get_type_elem(elem):
    try:
        return elem.Document.GetElement(elem.GetTypeId())
    except:
        return None


def get_type_name(elem):
    return get_element_name(get_type_elem(elem))


def get_family_name(elem):
    type_elem = get_type_elem(elem)

    try:
        if hasattr(type_elem, "FamilyName"):
            return clean_text(type_elem.FamilyName)
    except:
        pass

    try:
        if hasattr(elem, "Symbol") and elem.Symbol and elem.Symbol.Family:
            return get_element_name(elem.Symbol.Family)
    except:
        pass

    return ""


def get_first_mep_system(elem):
    """
    Lay MEPSystem tu element.

    Ghi chu V4.3:
    - Pipe Fitting / Pipe Accessory nhieu khi khong tra ve System Name bang BuiltInParameter.
    - Khi do doc qua Connector.MEPSystem se on dinh hon.
    - Ham nay chi doc du lieu, khong thay doi cau truc so sanh rule.
    """
    if elem is None:
        return None

    try:
        sys = elem.MEPSystem
        if sys:
            return sys
    except:
        pass

    try:
        mep_model = elem.MEPModel
        if mep_model is not None:
            cm = mep_model.ConnectorManager
            if cm is not None:
                for connector in cm.Connectors:
                    try:
                        sys = connector.MEPSystem
                        if sys:
                            return sys
                    except:
                        pass
    except:
        pass

    try:
        cm = elem.ConnectorManager
        if cm is not None:
            for connector in cm.Connectors:
                try:
                    sys = connector.MEPSystem
                    if sys:
                        return sys
                except:
                    pass
    except:
        pass

    return None


def get_system_param_from_connector(elem, param_names, bip_names=None):
    sys = get_first_mep_system(elem)
    if sys is None:
        return ""

    if bip_names:
        value = get_bip_param_text(sys, bip_names)
        if value:
            return value

    for name in param_names:
        value = get_param_as_text(sys, name)
        if value:
            return value

    return ""


def get_system_name(elem):
    value = get_bip_param_text(elem, ["RBS_SYSTEM_NAME_PARAM"])
    if value:
        return value

    value = get_param_as_text(elem, "System Name")
    if value:
        return value

    sys = get_first_mep_system(elem)
    if sys:
        return get_element_name(sys)

    return ""


def get_system_abbreviation(elem):
    value = get_bip_param_text(elem, ["RBS_DUCT_PIPE_SYSTEM_ABBREVIATION_PARAM"])
    if value:
        return value

    value = get_param_as_text(elem, "System Abbreviation")
    if value:
        return value

    value = get_system_param_from_connector(
        elem,
        ["System Abbreviation", "Abbreviation", "System Abbrev."],
        ["RBS_DUCT_PIPE_SYSTEM_ABBREVIATION_PARAM"]
    )
    if value:
        return value

    return ""


def get_system_classification(elem):
    value = get_bip_param_text(elem, ["RBS_SYSTEM_CLASSIFICATION_PARAM"])
    if value:
        return value

    value = get_param_as_text(elem, "System Classification")
    if value:
        return value

    value = get_system_param_from_connector(
        elem,
        ["System Classification", "Classification"],
        ["RBS_SYSTEM_CLASSIFICATION_PARAM"]
    )
    if value:
        return value

    return ""


def get_level_name(elem):
    try:
        level = elem.Document.GetElement(elem.LevelId)
        if level:
            return get_element_name(level)
    except:
        pass

    value = get_bip_param_text(elem, ["RBS_START_LEVEL_PARAM", "FAMILY_LEVEL_PARAM", "LEVEL_PARAM", "SCHEDULE_LEVEL_PARAM"])
    if value:
        return value

    for pname in ["Reference Level", "Level", "Schedule Level", "Base Level"]:
        value = get_param_as_text(elem, pname)
        if value:
            return value

    return ""


def get_workset_name(elem):
    try:
        ws = elem.Document.GetWorksetTable().GetWorkset(elem.WorksetId)
        if ws:
            return clean_text(ws.Name)
    except:
        pass
    return ""


def get_phase_name(elem):
    try:
        p = elem.get_Parameter(BuiltInParameter.PHASE_CREATED)
        if p and p.HasValue:
            phase = elem.Document.GetElement(p.AsElementId())
            if phase:
                return get_element_name(phase)
    except:
        pass
    return ""


def get_connector_sizes_mm(elem):
    """
    Doc size tu connector cho Pipe Fitting / Accessory / Flex.

    Revit family fitting thuong khong co RBS_PIPE_DIAMETER_PARAM nhu Pipe curve.
    Connector Radius/Width/Height giup lay duoc Diameter/Width/Height de mapping size khong bi fail.
    """
    result = {
        "diameters": [],
        "widths": [],
        "heights": []
    }

    managers = []

    try:
        if elem.MEPModel is not None and elem.MEPModel.ConnectorManager is not None:
            managers.append(elem.MEPModel.ConnectorManager)
    except:
        pass

    try:
        if elem.ConnectorManager is not None:
            managers.append(elem.ConnectorManager)
    except:
        pass

    for cm in managers:
        try:
            for c in cm.Connectors:
                try:
                    r = c.Radius
                    if r and r > 0:
                        result["diameters"].append(length_internal_to_mm(r * 2.0))
                except:
                    pass

                try:
                    w = c.Width
                    if w and w > 0:
                        result["widths"].append(length_internal_to_mm(w))
                except:
                    pass

                try:
                    h = c.Height
                    if h and h > 0:
                        result["heights"].append(length_internal_to_mm(h))
                except:
                    pass
        except:
            pass

    return result


def get_connector_primary_diameter_mm(elem):
    sizes = get_connector_sizes_mm(elem)
    vals = sizes.get("diameters", [])
    if len(vals) > 0:
        return max(vals)
    return None


def get_connector_primary_width_mm(elem):
    sizes = get_connector_sizes_mm(elem)
    vals = sizes.get("widths", [])
    if len(vals) > 0:
        return max(vals)
    return None


def get_connector_primary_height_mm(elem):
    sizes = get_connector_sizes_mm(elem)
    vals = sizes.get("heights", [])
    if len(vals) > 0:
        return max(vals)
    return None


def get_connector_overall_size_text(elem):
    sizes = get_connector_sizes_mm(elem)
    ds = []
    for d in sizes.get("diameters", []):
        txt = format_number(d)
        if txt not in ds:
            ds.append(txt)

    ws = []
    for w in sizes.get("widths", []):
        txt = format_number(w)
        if txt not in ws:
            ws.append(txt)

    hs = []
    for h in sizes.get("heights", []):
        txt = format_number(h)
        if txt not in hs:
            hs.append(txt)

    if len(ds) > 0:
        # Pipe fitting reducer thuong co 2 duong kinh. Giu text SizeText de Equals/Dropdown on dinh.
        return "-".join([x + " mm" for x in ds])

    if len(ws) > 0 and len(hs) > 0:
        return "-".join([ws[0] + "x" + hs[0] + " mm"])

    return ""


def get_diameter_mm(elem):
    value = get_length_param_mm(
        elem,
        ["RBS_PIPE_DIAMETER_PARAM", "RBS_CURVE_DIAMETER_PARAM", "CONNECTOR_DIAMETER"],
        ["Diameter", "DN", "Nominal Diameter", "Nominal diameter", "Size", "Connector Diameter"]
    )
    if value is not None:
        return value

    return get_connector_primary_diameter_mm(elem)


def get_width_mm(elem):
    value = get_length_param_mm(elem, ["RBS_CURVE_WIDTH_PARAM", "CONNECTOR_WIDTH"], ["Width", "Nominal Width", "Connector Width"])
    if value is not None:
        return value

    return get_connector_primary_width_mm(elem)


def get_height_mm(elem):
    value = get_length_param_mm(elem, ["RBS_CURVE_HEIGHT_PARAM", "CONNECTOR_HEIGHT"], ["Height", "Nominal Height", "Connector Height"])
    if value is not None:
        return value

    return get_connector_primary_height_mm(elem)


def get_length_mm(elem):
    return get_length_param_mm(elem, ["CURVE_ELEM_LENGTH"], ["Length"])


def is_insulation_element(elem):
    """
    Nhận diện element insulation.

    Ghi chú:
    - Dùng class name để tương thích nhiều version Revit.
    - Không ảnh hưởng cấu trúc so sánh rule, chỉ phục vụ sửa insulation hiện có.
    """
    if elem is None:
        return False

    try:
        cname = elem.GetType().Name
        if "Insulation" in cname:
            return True
    except:
        pass

    try:
        if elem.Category and "Insulation" in elem.Category.Name:
            return True
    except:
        pass

    return False


def get_existing_insulation_elements(elem):
    """
    Lấy danh sách insulation đang bám vào element MEP.

    Dùng GetDependentElements(None) vì Pipe/Duct/Fitting thường lưu insulation
    như dependent element của chính host element.
    """
    result = []

    try:
        dep_ids = elem.GetDependentElements(None)
    except:
        dep_ids = []

    for dep_id in dep_ids:
        try:
            dep = elem.Document.GetElement(dep_id)
            if is_insulation_element(dep):
                result.append(dep)
        except:
            pass

    return result


def get_insulation_existing(elem):
    """
    Giá trị hiển thị trong tab Active View Elements.
    """
    return "YES" if len(get_existing_insulation_elements(elem)) > 0 else "NO"


def get_insulation_type_name(insulation_elem):
    """
    Lấy tên type của insulation hiện có.
    """
    try:
        type_elem = insulation_elem.Document.GetElement(insulation_elem.GetTypeId())
        return get_element_name(type_elem)
    except:
        return ""


def get_insulation_thickness_mm(insulation_elem):
    """
    Lấy thickness của insulation hiện có, đơn vị mm.

    Revit API có thể khác nhẹ giữa version / loại pipe-duct insulation,
    nên hàm này thử nhiều đường đọc:
    1. Property Thickness nếu class có hỗ trợ.
    2. BuiltInParameter nếu tồn tại.
    3. LookupParameter theo các tên thường gặp.
    """
    if insulation_elem is None:
        return None

    # Cách 1: property Thickness, thường trả về internal feet.
    try:
        value = insulation_elem.Thickness
        if value is not None:
            return length_internal_to_mm(value)
    except:
        pass

    # Cách 2: built-in parameter. Dùng getattr để tránh lỗi khác version Revit.
    bip_names = [
        "RBS_PIPE_INSULATION_THICKNESS",
        "RBS_DUCT_INSULATION_THICKNESS",
        "RBS_INSULATION_THICKNESS",
        "RBS_REFERENCE_INSULATION_THICKNESS",
        "CURVE_ELEM_INSULATION_THICKNESS"
    ]

    for bip_name in bip_names:
        try:
            bip = getattr(BuiltInParameter, bip_name)
            p = insulation_elem.get_Parameter(bip)
            if p and p.HasValue:
                if p.StorageType == StorageType.Double:
                    return length_internal_to_mm(p.AsDouble())
                val = try_float(p.AsValueString())
                if val is not None:
                    return val
        except:
            pass

    # Cách 3: lookup theo tên parameter thường gặp, có thêm tiếng Đức để phòng model EU.
    param_names = [
        "Thickness",
        "Insulation Thickness",
        "Insulation thickness",
        "Dicke",
        "Dämmstärke",
        "Isolierdicke"
    ]

    for name in param_names:
        try:
            p = insulation_elem.LookupParameter(name)
            if p and p.HasValue:
                if p.StorageType == StorageType.Double:
                    return length_internal_to_mm(p.AsDouble())
                val = try_float(p.AsValueString())
                if val is not None:
                    return val
        except:
            pass

    return None


def insulation_matches_rule(existing_insulations, target_type, target_thickness_mm):
    """
    Kiem tra insulation hien co co dung theo rule khong.

    V4.5:
    - Khong bat buoc chi co 1 insulation.
    - Neu tat ca insulation hien co dung type + thickness thi skip.
    - Neu co bat ky insulation nao sai type/thickness thi repair bang cach EDIT,
      khong xoa, de giu ElementId cua insulation.
    """
    if len(existing_insulations) == 0:
        return False, "No existing insulation"

    target_type_name = get_element_name(target_type)
    problems = []

    for ins in existing_insulations:
        ins_id = ""
        try:
            ins_id = element_id_value(ins.Id)
        except:
            ins_id = "?"

        current_type_name = get_insulation_type_name(ins)
        if not strings_equal(current_type_name, target_type_name):
            problems.append(
                "Insulation {0}: type mismatch current='{1}', target='{2}'".format(
                    ins_id,
                    current_type_name,
                    target_type_name
                )
            )

        current_thickness_mm = get_insulation_thickness_mm(ins)
        if current_thickness_mm is None:
            problems.append("Insulation {0}: cannot read thickness".format(ins_id))
        elif abs(float(current_thickness_mm) - float(target_thickness_mm)) > INSULATION_THICKNESS_TOLERANCE_MM:
            problems.append(
                "Insulation {0}: thickness mismatch current='{1} mm', target='{2} mm'".format(
                    ins_id,
                    format_number(current_thickness_mm),
                    format_number(target_thickness_mm)
                )
            )

    if len(problems) > 0:
        return False, "; ".join(problems)

    return True, "Existing insulation already matches rule"


def get_field_value(elem, field_name, active_view=None):
    """
    Hàm trung tâm để đọc giá trị field/parameter từ element.
    Muốn thêm FieldName mới thì thêm case tại đây.
    Nếu không có case riêng, tool sẽ fallback sang LookupParameter(field_name).
    """
    name = clean_text(field_name)

    if name == "":
        return ""

    if name == "System Name":
        return get_system_name(elem)

    if name == "System Abbreviation":
        return get_system_abbreviation(elem)

    if name == "System Classification":
        return get_system_classification(elem)

    if name == "Level Name":
        return get_level_name(elem)

    if name == "Workset":
        return get_workset_name(elem)

    if name == "Phase":
        return get_phase_name(elem)

    if name == "View Name":
        if active_view:
            return get_element_name(active_view)
        return ""

    if name == "Category":
        return get_category_name(elem)

    if name == "Family Name":
        return get_family_name(elem)

    if name == "Type Name":
        return get_type_name(elem)

    if name == "Diameter":
        return format_number(get_diameter_mm(elem))

    if name == "Width":
        return format_number(get_width_mm(elem))

    if name == "Height":
        return format_number(get_height_mm(elem))

    if name == "Length":
        return format_number(get_length_mm(elem))

    if name == "Overall Size":
        # Pipe fitting/accessory co the khong co Overall Size truc tiep.
        # Uu tien parameter text cua family, sau do fallback sang connector size.
        for pname in ["Overall Size", "Size", "Nominal Size", "Connector Size"]:
            value = get_param_as_text(elem, pname)
            if value:
                return value

        value = get_connector_overall_size_text(elem)
        if value:
            return value

        return ""

    if name == "Comments":
        return get_param_as_text(elem, "Comments")

    if name == "Mark":
        return get_param_as_text(elem, "Mark")

    if name == "Type Comments":
        type_elem = get_type_elem(elem)
        return get_param_as_text(type_elem, "Type Comments")

    if name == "Installation Area":
        value = get_param_as_text(elem, "Installation Area")
        if value:
            return value
        return get_param_as_text(elem, "SI_InstallationArea")

    if name == "Service Type":
        return get_param_as_text(elem, "Service Type")

    if name == "Insulation Area":
        return get_param_as_text(elem, "Insulation Area")

    if name == "InsulationTypeName":
        return get_param_as_text(elem, "InsulationTypeName")

    return get_param_as_text(elem, name)


# =============================================================================
# COLLECTORS
# =============================================================================

def builtin_category_by_name(name):
    try:
        return getattr(BuiltInCategory, name)
    except:
        return None


def builtin_category_int(bic):
    """
    Lấy integer value của BuiltInCategory.

    Ghi chú:
    - Revit 2023 thường dùng IntegerValue qua ElementId.
    - BuiltInCategory là enum nên IronPython có thể ép int trực tiếp.
    - Hàm này giúp nhận diện category không phụ thuộc ngôn ngữ UI của Revit.
    """
    try:
        return int(bic)
    except:
        pass

    try:
        return bic.value__
    except:
        return None


def category_id_int(category):
    if category is None:
        return None

    try:
        return category.Id.IntegerValue
    except:
        pass

    try:
        return category.Id.Value
    except:
        return None


def element_category_is(elem, bic_name):
    """
    Kiểm tra category bằng BuiltInCategory thay vì Category.Name.
    Cách này không bị lỗi khi Revit dùng tiếng Đức/tiếng Việt.
    """
    try:
        bic = builtin_category_by_name(bic_name)
        if bic is None:
            return False

        elem_cat_id = category_id_int(elem.Category)
        bic_id = builtin_category_int(bic)

        return elem_cat_id is not None and bic_id is not None and elem_cat_id == bic_id
    except:
        return False


def get_categories_for_element_filter(element_filter):
    f = clean_text(element_filter)

    # Có thêm vài alias đề phòng khác biệt tên enum theo version Revit.
    pipe_names = [
        "OST_PipeCurves",
        "OST_PipeFitting",
        "OST_PipeFittings",
        "OST_PipeAccessory",
        "OST_PipeAccessories",
        "OST_FlexPipeCurves"
    ]
    duct_names = [
        "OST_DuctCurves",
        "OST_DuctFitting",
        "OST_DuctFittings",
        "OST_DuctAccessory",
        "OST_DuctAccessories",
        "OST_FlexDuctCurves"
    ]

    if f == "Pipe":
        names = ["OST_PipeCurves"]
    elif f == "Pipe Fitting":
        names = ["OST_PipeFitting"]
    elif f == "Pipe Accessory":
        names = ["OST_PipeAccessory"]
    elif f == "Flex Pipe":
        names = ["OST_FlexPipeCurves"]
    elif f == "Duct":
        names = ["OST_DuctCurves"]
    elif f == "Duct Fitting":
        names = ["OST_DuctFitting"]
    elif f == "Duct Accessory":
        names = ["OST_DuctAccessory"]
    elif f == "Flex Duct":
        names = ["OST_FlexDuctCurves"]
    elif f == "All Pipe":
        names = pipe_names
    elif f == "All Duct":
        names = duct_names
    elif f == "Both":
        names = ["OST_PipeCurves", "OST_DuctCurves"]
    elif f == "All":
        names = pipe_names + duct_names
    else:
        names = ["OST_PipeCurves", "OST_DuctCurves"]

    result = []
    for name in names:
        bic = builtin_category_by_name(name)
        if bic is not None:
            result.append(bic)
    return result


def get_categories_for_dropdown_scan():
    """
    Category dùng riêng để tạo datasource dropdown cho Rule DataGrid.

    Lý do sửa V4.1:
    - Trước đây dropdown bị phụ thuộc Element Filter.
    - Nếu Element Filter = Both thì chỉ quét Pipe/Duct chính, không quét fitting/accessory/flex.
    - Kết quả: Rule ElementType = Pipe Fitting + RevitFieldName = Family Name bị trắng.

    Hàm này luôn quét tất cả nhóm element có thể add insulation trong Active View
    để tạo bucket dropdown theo ElementType. UI preview vẫn dùng Element Filter như cũ.
    """
    return get_categories_for_element_filter("All")


def collect_active_view_elements_by_categories(revit_doc, active_view, cats):
    elements = []

    for bic in cats:
        try:
            collector = (
                FilteredElementCollector(revit_doc, active_view.Id)
                .OfCategory(bic)
                .WhereElementIsNotElementType()
            )

            for elem in collector:
                elements.append(elem)
        except:
            pass

    return elements


def collect_active_view_pipe_ducts(revit_doc, active_view, element_filter):
    cats = get_categories_for_element_filter(element_filter)
    return collect_active_view_elements_by_categories(revit_doc, active_view, cats)


def collect_selected_insulation_elements(revit_doc, uidoc_obj, element_filter):
    elements = []
    try:
        ids = list(uidoc_obj.Selection.GetElementIds())
    except:
        ids = []

    for eid in ids:
        try:
            elem = revit_doc.GetElement(eid)
            if elem is None:
                continue
            elem_type = get_element_basic_type(elem)
            if element_type_match(element_filter, elem_type):
                elements.append(elem)
        except:
            pass

    return elements



def get_enabled_rule_element_filters(rules):
    """
    Lay danh sach ElementType can scan khi Apply Rules.

    V4.3:
    - Truoc day Apply Rules dung Element Filter cua UI preview.
    - Neu Element Filter dang la Pipe/Both thi Pipe Fitting khong duoc quet, nen rule Pipe Fitting khong co co hoi match.
    - Ham nay dung chinh Rule.ElementType de quyet dinh element nao can kiem tra.
    - Khong thay doi cau truc so sanh, chi sua tap element dau vao.
    """
    filters = []

    for rule in rules:
        try:
            if not rule.Enabled:
                continue
        except:
            pass

        f = clean_text(getattr(rule, "ElementType", ""))
        if f == "":
            f = "All"

        if f not in filters:
            filters.append(f)

        if f == "All":
            return ["All"]

    if len(filters) == 0:
        filters.append("All")

    return filters


def get_categories_for_rule_filters(rule_filters):
    cats = []
    seen = {}

    for f in rule_filters:
        for bic in get_categories_for_element_filter(f):
            try:
                key = int(bic)
            except:
                key = clean_text(bic)
            if key in seen:
                continue
            seen[key] = True
            cats.append(bic)

    return cats


def element_matches_any_rule_element_type(elem, rules):
    elem_type = get_element_basic_type(elem)

    for rule in rules:
        try:
            if not rule.Enabled:
                continue
        except:
            pass

        rule_type = clean_text(getattr(rule, "ElementType", ""))
        if element_type_match(rule_type, elem_type):
            return True

    return False


def collect_apply_scope_elements(revit_doc, active_view, uidoc_obj, scope, rules):
    """
    Collect element cho Apply Rules.

    Active View:
    - Quet category dua tren Rule.ElementType, khong dua vao Element Filter preview.

    Selection:
    - Lay element dang chon trong Revit, sau do loc theo Rule.ElementType.
    """
    if scope == "Selection":
        elems = []
        try:
            ids = list(uidoc_obj.Selection.GetElementIds())
        except:
            ids = []

        for eid in ids:
            try:
                elem = revit_doc.GetElement(eid)
                if elem is None:
                    continue
                if element_matches_any_rule_element_type(elem, rules):
                    elems.append(elem)
            except:
                pass

        return elems

    rule_filters = get_enabled_rule_element_filters(rules)
    cats = get_categories_for_rule_filters(rule_filters)
    return collect_active_view_elements_by_categories(revit_doc, active_view, cats)


def collect_parameter_names_from_element(elem, target):
    try:
        for p in elem.Parameters:
            try:
                add_unique_text(target, p.Definition.Name)
            except:
                pass
    except:
        pass

    try:
        type_elem = elem.Document.GetElement(elem.GetTypeId())
        for p in type_elem.Parameters:
            try:
                add_unique_text(target, p.Definition.Name)
            except:
                pass
    except:
        pass


def make_preview_item(elem, active_view, extra_fields):
    item = ElementPreviewItem()

    item.ElementId = clean_text(element_id_value(elem.Id))
    item.Category = get_category_name(elem)
    item.SystemName = get_system_name(elem)
    item.SystemAbbreviation = get_system_abbreviation(elem)
    item.SystemClassification = get_system_classification(elem)
    item.LevelName = get_level_name(elem)
    item.Workset = get_workset_name(elem)
    item.Phase = get_phase_name(elem)
    item.FamilyName = get_family_name(elem)
    item.TypeName = get_type_name(elem)
    item.Diameter = format_number(get_diameter_mm(elem))
    item.Width = format_number(get_width_mm(elem))
    item.Height = format_number(get_height_mm(elem))
    item.Length = format_number(get_length_mm(elem))
    item.OverallSize = get_field_value(elem, "Overall Size", active_view)
    item.InsulationExisting = get_insulation_existing(elem)
    item.Comments = get_param_as_text(elem, "Comments")
    item.Mark = get_param_as_text(elem, "Mark")

    for field_name in extra_fields:
        binding = safe_binding_name(field_name)
        setattr(item, binding, get_field_value(elem, field_name, active_view))

    return item


def make_dropdown_data(include_by_type=True):
    """
    Tạo dictionary chứa giá trị dropdown.

    V4.0:
    - Giữ dropdown global như V3.3/V3.9.
    - Bổ sung bucket theo ElementType để Rule DataGrid lọc đúng theo từng dòng rule.
      Ví dụ: Rule.ElementType = Pipe Fitting, RevitFieldName = Family Name
      -> dropdown chỉ lấy Family Name của Pipe Fitting.
    """
    data = {}

    for field in BASE_REVIT_FIELD_OPTIONS:
        field = clean_text(field)
        if field != "":
            data[field] = {}

    for field in [
        "System Name",
        "System Abbreviation",
        "System Classification",
        "Level Name",
        "Workset",
        "Phase",
        "View Name",
        "Category",
        "Family Name",
        "Type Name",
        "InsulationTypeName",
        "Diameter",
        "Width",
        "Height",
        "Length",
        "Overall Size",
        "Room Name",
        "Space Name",
        "Installation Area",
        "Service Type",
        "Insulation Area",
        "SI_InstallationArea",
        "TGA_SystemCode",
        "Comments",
        "Mark",
        "Type Comments"
    ]:
        if field not in data:
            data[field] = {}

    data["ParameterNames"] = {}

    if include_by_type:
        data[DROPDOWN_BY_ELEMENT_TYPE_KEY] = {}

    return data


def ensure_dropdown_type_bucket(data, elem_type):
    """
    Tạo bucket dropdown theo ElementType.
    Bucket này dùng riêng cho UI, không ảnh hưởng logic so sánh rule.
    """
    elem_type = clean_text(elem_type)
    if elem_type == "":
        elem_type = "Unknown"

    if DROPDOWN_BY_ELEMENT_TYPE_KEY not in data:
        data[DROPDOWN_BY_ELEMENT_TYPE_KEY] = {}

    by_type = data[DROPDOWN_BY_ELEMENT_TYPE_KEY]

    if elem_type not in by_type:
        by_type[elem_type] = make_dropdown_data(include_by_type=False)

    return by_type[elem_type]


def add_dropdown_value(data, field_name, value):
    field_name = clean_text(field_name)
    if field_name == "":
        return

    if field_name not in data:
        data[field_name] = {}

    add_unique_text(data[field_name], value)


def matching_element_type_keys(rule_element_type):
    """
    Chuyển ElementType của rule thành các bucket element type cần dùng để lấy dropdown.
    Hàm này chỉ phục vụ UI dropdown, không thay đổi logic so sánh.
    """
    t = clean_text(rule_element_type)

    pipe_types = ["Pipe", "Pipe Fitting", "Pipe Accessory", "Flex Pipe"]
    duct_types = ["Duct", "Duct Fitting", "Duct Accessory", "Flex Duct"]

    if t == "All Pipe":
        return pipe_types

    if t == "All Duct":
        return duct_types

    if t == "Both":
        return ["Pipe", "Duct"]

    if t == "All" or t == "":
        return pipe_types + duct_types

    return [t]

def collect_parameter_names_and_values_from_element(elem, active_view, data):
    """
    Lấy tên parameter và giá trị parameter thật để làm dropdown cho Rule DataGrid.

    Ví dụ:
    - RevitFieldName = SI_InstallationArea
    - Element có SI_InstallationArea = Technik
    -> Rule DataGrid cột tương ứng sẽ có dropdown value "Technik".

    Hàm này đọc cả instance parameter và type parameter.
    """
    try:
        for p in elem.Parameters:
            try:
                pname = clean_text(p.Definition.Name)
                if pname == "":
                    continue

                add_unique_text(data["ParameterNames"], pname)

                add_dropdown_value(data, pname, get_param_as_text(elem, pname))
            except:
                pass
    except:
        pass

    try:
        type_elem = elem.Document.GetElement(elem.GetTypeId())
        if type_elem:
            for p in type_elem.Parameters:
                try:
                    pname = clean_text(p.Definition.Name)
                    if pname == "":
                        continue

                    add_unique_text(data["ParameterNames"], pname)

                    add_dropdown_value(data, pname, get_param_as_text(type_elem, pname))
                except:
                    pass
    except:
        pass


def add_insulation_type_to_ui_buckets(data, elem_group, type_name):
    """
    Đưa insulation type vào đúng nhóm Pipe/Duct cho dropdown UI.
    Pipe Fitting/Pipe Accessory/Flex Pipe vẫn dùng Pipe insulation type.
    Duct Fitting/Duct Accessory/Flex Duct vẫn dùng Duct insulation type.
    """
    type_name = clean_text(type_name)
    if type_name == "":
        return

    add_dropdown_value(data, "InsulationTypeName", type_name)

    if elem_group == "Pipe":
        target_types = ["Pipe", "Pipe Fitting", "Pipe Accessory", "Flex Pipe"]
    else:
        target_types = ["Duct", "Duct Fitting", "Duct Accessory", "Flex Duct"]

    for t in target_types:
        bucket = ensure_dropdown_type_bucket(data, t)
        add_dropdown_value(bucket, "InsulationTypeName", type_name)


def collect_insulation_type_names(revit_doc, data):
    # Cách 1: collect bằng API class nếu có.
    if PipeInsulationType is not None:
        try:
            for itype in FilteredElementCollector(revit_doc).OfClass(PipeInsulationType):
                add_insulation_type_to_ui_buckets(data, "Pipe", get_element_name(itype))
        except:
            pass

    if DuctInsulationType is not None:
        try:
            for itype in FilteredElementCollector(revit_doc).OfClass(DuctInsulationType):
                add_insulation_type_to_ui_buckets(data, "Duct", get_element_name(itype))
        except:
            pass

    # Cách 2: fallback bằng category nếu API class không có.
    for bic_name, elem_group in [
        ("OST_PipeInsulations", "Pipe"),
        ("OST_DuctInsulations", "Duct")
    ]:
        try:
            bic = getattr(BuiltInCategory, bic_name)
            for itype in FilteredElementCollector(revit_doc).OfCategory(bic).WhereElementIsElementType():
                add_insulation_type_to_ui_buckets(data, elem_group, get_element_name(itype))
        except:
            pass

def collect_dropdown_data(revit_doc, active_view, element_filter):
    """
    Collect dropdown values từ Revit.

    BẢN V4.1:
    - Giữ cơ chế mở nhanh: chỉ chạy khi người dùng bấm refresh, không chạy lúc startup.
    - Không thay đổi cấu trúc so sánh rule.
    - Datasource dropdown Rule DataGrid được quét từ tất cả nhóm element có thể add insulation
      trong Active View: Pipe, Pipe Fitting, Pipe Accessory, Flex Pipe, Duct, Duct Fitting,
      Duct Accessory, Flex Duct.
    - Vì vậy Rule ElementType = Pipe Fitting sẽ có bucket riêng cho Family Name, Type Name,
      System Name, Overall Size, custom parameter... nếu trong Active View có Pipe Fitting.
    - Element Filter vẫn chỉ dùng cho tab preview và Apply Scope, không làm nghèo dropdown mapping.
    """
    data = make_dropdown_data()

    # Quét dropdown bằng tất cả category có thể add insulation.
    # Không phụ thuộc Element Filter, để Pipe Fitting / Accessory / Flex không bị trắng dropdown.
    dropdown_cats = get_categories_for_dropdown_scan()
    elems = collect_active_view_elements_by_categories(revit_doc, active_view, dropdown_cats)

    scan_count = 0
    for elem in elems:
        if MAX_DROPDOWN_ELEMENT_SCAN is not None and scan_count >= MAX_DROPDOWN_ELEMENT_SCAN:
            break
        scan_count += 1

        elem_type = get_element_basic_type(elem)
        type_bucket = ensure_dropdown_type_bucket(data, elem_type)

        # Quét tất cả field đã biết, không chỉ System Name.
        # Đồng thời ghi vào global dropdown và dropdown riêng theo ElementType.
        for field in list(data.keys()):
            if field in ["ParameterNames", DROPDOWN_BY_ELEMENT_TYPE_KEY]:
                continue

            value = get_field_value(elem, field, active_view)
            add_dropdown_value(data, field, value)
            add_dropdown_value(type_bucket, field, value)

        collect_parameter_names_and_values_from_element(elem, active_view, data)
        collect_parameter_names_and_values_from_element(elem, active_view, type_bucket)

    # Các danh sách project-level nhẹ, vẫn giữ để dropdown có dữ liệu dù active view chưa có element tương ứng.
    try:
        for level in FilteredElementCollector(revit_doc).OfClass(Level):
            add_unique_text(data["Level Name"], get_element_name(level))
    except:
        pass

    try:
        for ws in FilteredWorksetCollector(revit_doc).OfKind(WorksetKind.UserWorkset):
            add_unique_text(data["Workset"], ws.Name)
    except:
        pass

    try:
        for phase in FilteredElementCollector(revit_doc).OfClass(Phase):
            add_unique_text(data["Phase"], get_element_name(phase))
    except:
        pass

    try:
        # Để mở tool và refresh dropdown mượt hơn, mặc định chỉ thêm Active View hiện tại.
        # Nếu muốn quét toàn bộ View trong project, đổi INCLUDE_ALL_VIEWS_IN_DROPDOWN = True.
        if active_view:
            add_unique_text(data["View Name"], get_element_name(active_view))

        if INCLUDE_ALL_VIEWS_IN_DROPDOWN:
            for v in FilteredElementCollector(revit_doc).OfClass(View):
                if not v.IsTemplate:
                    add_unique_text(data["View Name"], get_element_name(v))
    except:
        pass

    collect_insulation_type_names(revit_doc, data)

    return data


def get_element_basic_type(elem):
    """
    Trả về ElementType chuẩn dùng cho Rule Editor.

    Sửa V4.1:
    - Ưu tiên nhận diện bằng BuiltInCategory.Id để không phụ thuộc ngôn ngữ Category.Name.
    - Nếu dùng Category.Name trên Revit tiếng Đức, Pipe Fitting có thể thành tên bản địa,
      làm bucket dropdown lưu sai key và Rule DataGrid bị trắng.
    """
    if element_category_is(elem, "OST_PipeFitting") or element_category_is(elem, "OST_PipeFittings"):
        return "Pipe Fitting"
    if element_category_is(elem, "OST_PipeAccessory") or element_category_is(elem, "OST_PipeAccessories"):
        return "Pipe Accessory"
    if element_category_is(elem, "OST_FlexPipeCurves"):
        return "Flex Pipe"
    if element_category_is(elem, "OST_PipeCurves"):
        return "Pipe"

    if element_category_is(elem, "OST_DuctFitting") or element_category_is(elem, "OST_DuctFittings"):
        return "Duct Fitting"
    if element_category_is(elem, "OST_DuctAccessory") or element_category_is(elem, "OST_DuctAccessories"):
        return "Duct Accessory"
    if element_category_is(elem, "OST_FlexDuctCurves"):
        return "Flex Duct"
    if element_category_is(elem, "OST_DuctCurves"):
        return "Duct"

    # Fallback cũ theo Category.Name, giữ để không phá các trường hợp đặc biệt.
    cat = get_category_name(elem).lower()

    if "pipe fittings" in cat or "pipe fitting" in cat:
        return "Pipe Fitting"
    if "pipe accessories" in cat or "pipe accessory" in cat:
        return "Pipe Accessory"
    if "flex pipes" in cat or "flex pipe" in cat:
        return "Flex Pipe"
    if "pipes" in cat or "pipe" in cat:
        return "Pipe"

    if "duct fittings" in cat or "duct fitting" in cat:
        return "Duct Fitting"
    if "duct accessories" in cat or "duct accessory" in cat:
        return "Duct Accessory"
    if "flex ducts" in cat or "flex duct" in cat:
        return "Flex Duct"
    if "ducts" in cat or "duct" in cat:
        return "Duct"

    return get_category_name(elem)


def is_pipe_like_type(elem_type):
    return clean_text(elem_type) in ["Pipe", "Pipe Fitting", "Pipe Accessory", "Flex Pipe"]


def is_duct_like_type(elem_type):
    return clean_text(elem_type) in ["Duct", "Duct Fitting", "Duct Accessory", "Flex Duct"]


def element_type_match(rule_type, elem_type):
    rule_type = clean_text(rule_type)
    elem_type = clean_text(elem_type)

    if rule_type == "" or rule_type == "All":
        return True
    if rule_type == "Both":
        return elem_type in ["Pipe", "Duct"]
    if rule_type == "All Pipe":
        return is_pipe_like_type(elem_type)
    if rule_type == "All Duct":
        return is_duct_like_type(elem_type)
    return strings_equal(rule_type, elem_type)


def get_size_value_mm(elem):
    diameter = get_diameter_mm(elem)
    if diameter is not None:
        return diameter

    width = get_width_mm(elem)
    height = get_height_mm(elem)

    if width is not None and height is not None:
        return max(width, height)
    if width is not None:
        return width
    if height is not None:
        return height

    value = try_float(get_field_value(elem, "Overall Size", None))
    if value is not None:
        return value

    return None


# =============================================================================
# RULE EVALUATION
# =============================================================================

def evaluate_condition(element_value, operator_name, rule_value, report=None, value_type="Auto"):
    """
    So sanh gia tri element voi gia tri trong rule.
    Numeric operator se tu tach so khoi chuoi co suffix, vi du 100 mmø-80 mmø -> 100.
    ValueType = Text/SizeText giu so sanh chuoi cho Equals/Contains.
    ValueType = Number/LengthMM uu tien so sanh so cho Equals/NotEquals.
    """
    actual = clean_text(element_value)
    expected = clean_text(rule_value)
    op = clean_text(operator_name)
    vt = clean_text(value_type)

    if op == "":
        return True

    if vt in ["Number", "LengthMM"] and op in ["Equals", "NotEquals"]:
        actual_num = try_float(actual)
        expected_num = try_float(expected)
        if actual_num is None or expected_num is None:
            if report is not None:
                report.append("Numeric equals failed. actual='{0}', expected='{1}'".format(actual, expected))
            return False
        same = abs(actual_num - expected_num) < 0.000001
        return same if op == "Equals" else not same

    if op == "IsEmpty":
        return actual == ""

    if op == "IsNotEmpty":
        return actual != ""

    if op == "Equals":
        return strings_equal(actual, expected)

    if op == "NotEquals":
        return not strings_equal(actual, expected)

    if op == "Contains":
        return contains_text(actual, expected)

    if op == "NotContains":
        return not contains_text(actual, expected)

    if op == "StartsWith":
        return starts_with_text(actual, expected)

    if op == "EndsWith":
        return ends_with_text(actual, expected)

    actual_num = try_float(actual)
    expected_num = try_float(expected)

    if actual_num is None or expected_num is None:
        if report is not None:
            report.append("Numeric compare failed. actual='{0}', expected='{1}', operator={2}".format(actual, expected, op))
        return False

    if op == "GreaterThan":
        return actual_num > expected_num

    if op == "GreaterOrEqual":
        return actual_num >= expected_num

    if op == "LessThan":
        return actual_num < expected_num

    if op == "LessOrEqual":
        return actual_num <= expected_num

    if report is not None:
        report.append("Unknown operator: {0}".format(op))

    return False


def rule_basic_match(elem, rule):
    """
    Kiểm tra điều kiện cơ bản của rule.

    V4.3:
    - MinDN / MaxDN đã được loại khỏi Rule DataGrid.
    - Basic match chỉ kiểm tra Enabled và ElementType.
    - Nếu cần lọc kích thước, khai báo cột động trong Parameter Column Mapping:
      RevitFieldName = Diameter / Width / Height / Overall Size
      Operator = GreaterOrEqual / LessOrEqual / Equals...
    """
    if not rule.Enabled:
        return False

    elem_type = get_element_basic_type(elem)
    rule_type = clean_text(rule.ElementType)

    if not element_type_match(rule_type, elem_type):
        return False

    return True


def rule_dynamic_match(elem, rule, mappings, active_view=None, report=None):
    """
    Duyệt toàn bộ mapping:
    - mapping disabled -> bỏ qua
    - thiếu column/field/operator -> bỏ qua
    - ô rule trống -> bỏ qua
    - có đủ dữ liệu -> so sánh
    """
    for mapping in mappings:
        if not mapping.Enabled:
            continue

        col_name = clean_text(mapping.RuleColumnName)
        revit_field = clean_text(mapping.RevitFieldName)
        operator_name = clean_text(mapping.Operator)

        if col_name == "" or revit_field == "" or operator_name == "":
            continue

        rule_value = rule.get_dynamic_value(mapping)

        if clean_text(rule_value) == "":
            continue

        element_value = get_field_value(elem, revit_field, active_view)
        ok = evaluate_condition(element_value, operator_name, rule_value, report, getattr(mapping, "ValueType", "Auto"))

        if report is not None:
            report.append(
                "Rule {0}, {1}: RevitField='{2}', ElementValue='{3}', Operator='{4}', RuleValue='{5}', Result={6}".format(
                    clean_text(rule.RuleId),
                    col_name,
                    revit_field,
                    element_value,
                    operator_name,
                    rule_value,
                    ok
                )
            )

        if not ok:
            return False

    return True


def find_matching_rule(elem, rules, mappings, active_view=None, report=None):
    sorted_rules = sorted(
        list(rules),
        key=lambda r: try_float(r.Priority) if try_float(r.Priority) is not None else 999999
    )

    for rule in sorted_rules:
        if not rule_basic_match(elem, rule):
            if report is not None:
                report.append("Rule {0}: basic match failed.".format(rule.RuleId))
            continue

        if not rule_dynamic_match(elem, rule, mappings, active_view, report):
            if report is not None:
                report.append("Rule {0}: dynamic match failed.".format(rule.RuleId))
            continue

        if report is not None:
            report.append("Rule {0}: matched.".format(rule.RuleId))

        return rule

    return None


# =============================================================================
# VALIDATION
# =============================================================================

def validate_mapping(mappings):
    errors = []
    names = {}

    row = 1

    for mapping in mappings:
        col = clean_text(mapping.RuleColumnName)
        field = clean_text(mapping.RevitFieldName)
        op = clean_text(mapping.Operator)

        if col == "":
            errors.append("Mapping row {0}: RuleColumnName is empty.".format(row))
        else:
            key = col.lower()
            if key in names:
                errors.append("Mapping row {0}: duplicate RuleColumnName '{1}'.".format(row, col))
            else:
                names[key] = True

            if key in FIXED_RULE_FIELD_SET:
                errors.append("Mapping row {0}: RuleColumnName '{1}' conflicts with fixed rule column.".format(row, col))

        if mapping.Enabled:
            if field == "":
                errors.append("Mapping row {0}: RevitFieldName is required when Enabled is TRUE.".format(row))

            if op == "":
                errors.append("Mapping row {0}: Operator is required when Enabled is TRUE.".format(row))

        if op and op not in OPERATOR_OPTIONS:
            errors.append("Mapping row {0}: Operator '{1}' is invalid.".format(row, op))

        if clean_text(mapping.ValueSource) not in VALUE_SOURCE_OPTIONS:
            errors.append("Mapping row {0}: ValueSource is invalid.".format(row))

        if clean_text(getattr(mapping, "ValueType", "Auto")) not in VALUE_TYPE_OPTIONS:
            errors.append("Mapping row {0}: ValueType is invalid.".format(row))

        row += 1

    return errors


def validate_rules(rules, mappings):
    errors = []
    ids = {}

    row = 1

    for rule in rules:
        rid = clean_text(rule.RuleId)

        if rid == "":
            errors.append("Rule row {0}: RuleId is empty.".format(row))
        elif rid in ids:
            errors.append("Rule row {0}: duplicate RuleId '{1}'.".format(row, rid))
        else:
            ids[rid] = True

        if clean_text(rule.ElementType) not in ELEMENT_TYPE_OPTIONS:
            errors.append("Rule {0}: ElementType must be one of supported insulation element types.".format(rid))

        for fname in ["ThicknessMM", "Priority"]:
            value = clean_text(getattr(rule, fname))
            if value != "" and try_float(value) is None:
                errors.append("Rule {0}: {1} must be numeric.".format(rid, fname))

        if clean_text(rule.ThicknessMM) == "":
            errors.append("Rule {0}: ThicknessMM is required for Add Insulation.".format(rid))

        if clean_text(rule.InsulationTypeName) == "":
            errors.append("Rule {0}: InsulationTypeName is required for Add Insulation.".format(rid))

        for mapping in mappings:
            if not mapping.Enabled:
                continue

            col = clean_text(mapping.RuleColumnName)
            field = clean_text(mapping.RevitFieldName)
            op = clean_text(mapping.Operator)

            if col == "" or field == "" or op == "":
                continue

            rule_value = rule.get_dynamic_value(mapping)

            # Ô parameter động trống là hợp lệ và sẽ bỏ qua.
            if clean_text(rule_value) == "":
                continue

            if op in NUMERIC_OPERATORS and try_float(rule_value) is None:
                errors.append("Rule {0}: column '{1}' value must be numeric for operator {2}.".format(rid, col, op))

            if field in NUMERIC_FIELDS and op not in ["IsEmpty", "IsNotEmpty"]:
                if try_float(rule_value) is None:
                    errors.append("Rule {0}: column '{1}' value must be numeric for field {2}.".format(rid, col, field))

        row += 1

    return errors


def validate_all(mappings, rules):
    errors = []
    errors.extend(validate_mapping(mappings))
    errors.extend(validate_rules(rules, mappings))
    return errors


# =============================================================================
# INSULATION
# =============================================================================

def get_insulation_type_by_name(revit_doc, elem_type, type_name):
    target = clean_text(type_name)

    if target == "":
        return None

    classes = []

    if is_pipe_like_type(elem_type) and PipeInsulationType is not None:
        classes.append(PipeInsulationType)

    if is_duct_like_type(elem_type) and DuctInsulationType is not None:
        classes.append(DuctInsulationType)

    for cls in classes:
        try:
            for itype in FilteredElementCollector(revit_doc).OfClass(cls):
                if strings_equal(get_element_name(itype), target):
                    return itype
        except:
            pass

    # fallback category
    cats = []
    if is_pipe_like_type(elem_type):
        cats.append("OST_PipeInsulations")
    if is_duct_like_type(elem_type):
        cats.append("OST_DuctInsulations")

    for cat_name in cats:
        try:
            bic = getattr(BuiltInCategory, cat_name)
            for itype in FilteredElementCollector(revit_doc).OfCategory(bic).WhereElementIsElementType():
                if strings_equal(get_element_name(itype), target):
                    return itype
        except:
            pass

    return None


def create_new_insulation_raw(revit_doc, elem, elem_type, insulation_type, thickness_internal):
    """
    Tạo insulation mới cho element.

    Hàm này chỉ làm việc tạo mới, không tự kiểm tra existing insulation.
    Dùng chung cho case create và repair.
    """
    if is_pipe_like_type(elem_type) and PipeInsulation is not None:
        PipeInsulation.Create(revit_doc, elem.Id, insulation_type.Id, thickness_internal)
        return True, ""

    if is_duct_like_type(elem_type) and DuctInsulation is not None:
        DuctInsulation.Create(revit_doc, elem.Id, insulation_type.Id, thickness_internal)
        return True, ""

    return False, "PipeInsulation or DuctInsulation API class not available."


def set_insulation_type(insulation_elem, target_type):
    """
    Doi type cua insulation hien co ma khong xoa element.

    Muc tieu:
    - Giu nguyen ElementId cua insulation.
    - Neu type dang dung thi khong lam gi.
    """
    if insulation_elem is None or target_type is None:
        return False, "Invalid insulation/type"

    try:
        if insulation_elem.GetTypeId() == target_type.Id:
            return True, "Type already correct"
    except:
        pass

    try:
        insulation_elem.ChangeTypeId(target_type.Id)
        return True, "Type changed"
    except Exception as ex:
        return False, "Cannot change insulation type: {0}".format(ex)


def set_insulation_thickness(insulation_elem, thickness_internal):
    """
    Sua thickness cua insulation hien co ma khong xoa element.

    Thu tu sua:
    1. Property Thickness neu Revit API cho set.
    2. BuiltInParameter thickness neu ton tai.
    3. LookupParameter theo ten thuong gap.
    """
    if insulation_elem is None:
        return False, "Invalid insulation"

    # Cach 1: property Thickness.
    try:
        insulation_elem.Thickness = thickness_internal
        return True, "Thickness changed by property"
    except:
        pass

    # Cach 2: built-in parameter.
    bip_names = [
        "RBS_PIPE_INSULATION_THICKNESS",
        "RBS_DUCT_INSULATION_THICKNESS",
        "RBS_INSULATION_THICKNESS",
        "RBS_REFERENCE_INSULATION_THICKNESS",
        "CURVE_ELEM_INSULATION_THICKNESS"
    ]

    for bip_name in bip_names:
        try:
            bip = getattr(BuiltInParameter, bip_name)
            p = insulation_elem.get_Parameter(bip)
            if p and (not p.IsReadOnly):
                p.Set(thickness_internal)
                return True, "Thickness changed by BuiltInParameter {0}".format(bip_name)
        except:
            pass

    # Cach 3: lookup parameter.
    param_names = [
        "Thickness",
        "Insulation Thickness",
        "Insulation thickness",
        "Dicke",
        "Dämmstärke",
        "Isolierdicke"
    ]

    for name in param_names:
        try:
            p = insulation_elem.LookupParameter(name)
            if p and (not p.IsReadOnly):
                p.Set(thickness_internal)
                return True, "Thickness changed by parameter {0}".format(name)
        except:
            pass

    return False, "Cannot edit insulation thickness"


def edit_existing_insulations_by_rule(existing_insulations, insulation_type, thickness_internal):
    """
    Sua insulation hien co theo rule, khong xoa insulation cu.

    Quan trong:
    - Khong dung revit_doc.Delete().
    - Khong tao lai insulation neu da co insulation.
    - Muc tieu la giu nguyen ElementId cua insulation hien co.
    """
    edited = 0
    messages = []

    for ins in existing_insulations:
        ins_id = ""
        try:
            ins_id = element_id_value(ins.Id)
        except:
            ins_id = "?"

        ok_type, msg_type = set_insulation_type(ins, insulation_type)
        ok_thick, msg_thick = set_insulation_thickness(ins, thickness_internal)

        if ok_type and ok_thick:
            edited += 1
            messages.append("Insulation {0}: edited. {1}; {2}".format(ins_id, msg_type, msg_thick))
        else:
            messages.append("Insulation {0}: edit failed. {1}; {2}".format(ins_id, msg_type, msg_thick))

    if edited == len(existing_insulations):
        return True, "Edited {0} existing insulation element(s). {1}".format(edited, " | ".join(messages))

    return False, "Edited {0}/{1} insulation element(s). {2}".format(
        edited,
        len(existing_insulations),
        " | ".join(messages)
    )


def create_insulation_for_element(revit_doc, elem, rule):
    """
    Add / repair insulation cho element theo rule.

    Ket qua tra ve:
    - ok: True neu co tao moi hoac sua insulation.
    - message: noi dung report.
    - action: created / updated / skipped / failed.

    Logic V4.5:
    - Neu element chua co insulation -> tao moi.
    - Neu element da co insulation va dung type + thickness rule -> skip.
    - Neu element da co insulation nhung sai rule -> EDIT insulation hien co theo rule.
      Khong xoa insulation cu, khong tao lai, de giu ElementId cua insulation.
    """
    elem_type = get_element_basic_type(elem)
    thickness_mm = try_float(rule.ThicknessMM)

    if thickness_mm is None or thickness_mm <= 0:
        return False, "Invalid ThicknessMM", "failed"

    insulation_type = get_insulation_type_by_name(revit_doc, elem_type, rule.InsulationTypeName)

    if insulation_type is None:
        return False, "Insulation type not found: {0}".format(rule.InsulationTypeName), "failed"

    thickness_internal = mm_to_internal(thickness_mm)
    existing = get_existing_insulation_elements(elem)

    if len(existing) > 0:
        is_match, reason = insulation_matches_rule(existing, insulation_type, thickness_mm)

        if is_match:
            return False, "Existing insulation matches rule. Skipped.", "skipped"

        if not REPAIR_EXISTING_INSULATION_BY_RULE:
            return False, "Existing insulation does not match rule. Repair disabled. {0}".format(reason), "skipped"

        # Repair trong SubTransaction. Chi edit element insulation hien co,
        # khong xoa va khong tao lai, de khong doi ElementId.
        sub = SubTransaction(revit_doc)
        try:
            sub.Start()
            ok, msg = edit_existing_insulations_by_rule(existing, insulation_type, thickness_internal)

            if ok:
                sub.Commit()
                return True, "Edited existing insulation by rule. Reason: {0}. {1}".format(reason, msg), "updated"

            sub.RollBack()
            return False, "Edit failed and rolled back. Reason: {0}. {1}".format(reason, msg), "failed"

        except Exception as ex:
            try:
                sub.RollBack()
            except:
                pass
            return False, "Edit failed and rolled back: {0}".format(ex), "failed"

    try:
        ok, msg = create_new_insulation_raw(revit_doc, elem, elem_type, insulation_type, thickness_internal)
        if ok:
            return True, "Created new insulation.", "created"
        return False, msg, "failed"

    except Exception as ex:
        return False, str(ex), "failed"


# =============================================================================
# EXTERNAL EVENT
# =============================================================================

class RuleEditorExternalEventHandler(IExternalEventHandler):
    def __init__(self):
        self.window = None
        self.request = ""

    def GetName(self):
        return "Insulation Rule Editor ExternalEvent Handler"

    def Execute(self, uiapp):
        try:
            if self.window is None:
                return

            active_uidoc = uiapp.ActiveUIDocument
            revit_doc = active_uidoc.Document
            active_view = revit_doc.ActiveView

            if self.request == "REFRESH_ELEMENTS":
                self.window.external_refresh_elements(revit_doc, active_view)

            elif self.request == "REFRESH_DROPDOWNS":
                self.window.external_refresh_dropdowns(revit_doc, active_view)

            elif self.request == "APPLY_RULES":
                self.window.external_apply_rules(revit_doc, active_view, active_uidoc)

        except Exception as ex:
            output.print_md("# Rule Editor ExternalEvent Error")
            print(str(ex))
            print(traceback.format_exc())
            alert("ExternalEvent failed. See pyRevit output.")

        finally:
            self.request = ""


# =============================================================================
# WPF WINDOW
# =============================================================================

class RuleEditorWindow(forms.WPFWindow):
    def __init__(self, handler, external_event):
        forms.WPFWindow.__init__(self, XAML_FILE)

        self.handler = handler
        self.external_event = external_event

        # Khi đang mở tool và load CSV, tạm tắt các event rebuild để tránh UI bị ì.
        # Sau khi load xong, Rule DataGrid sẽ rebuild một lần duy nhất.
        self._is_loading_ui = True

        self.Mappings = ObservableCollection[object]()
        self.Rules = ObservableCollection[object]()
        self.FilteredRules = ObservableCollection[object]()
        self.ActiveElements = ObservableCollection[object]()
        self.FilteredActiveElements = ObservableCollection[object]()

        self.SelectedMapping = None

        self.OperatorOptions = ObservableCollection[object]()
        self.ValueSourceOptions = ObservableCollection[object]()
        self.ValueTypeOptions = ObservableCollection[object]()
        self.RevitFieldNameOptions = ObservableCollection[object]()
        self.ElementTypeOptions = ObservableCollection[object]()

        self.dropdown_data = make_dropdown_data()
        self._rule_filter_pending = False
        self._rule_options_refresh_pending = False

        self.mapping_csv_path = DEFAULT_MAPPING_CSV
        self.rule_csv_path = DEFAULT_RULE_CSV
        self._is_rebuilding_columns = False

        # Các cờ chống treo UI:
        # - Không rebuild DataGrid liên tục khi Mapping DataGrid phát nhiều event.
        # - Không CommitEdit trong lúc event đang bắn dây chuyền.
        # - Debounce rebuild bằng Dispatcher để UI có thời gian thở.
        self._suppress_mapping_events = True
        self._mapping_rebuild_pending = False
        self._mapping_dirty = False

        self.init_options()
        self.DataContext = self

        self.txtMappingCsvPath.Text = self.mapping_csv_path
        self.txtRuleCsvPath.Text = self.rule_csv_path

        self.bind_events()

        self.load_default_files_if_available()
        self.rebuild_rule_grid_columns()
        self.rebuild_active_element_columns([])

        # Bật lại event sau khi UI đã sẵn sàng.
        self._is_loading_ui = False
        self._suppress_mapping_events = False
        self.set_status("Ready. Dữ liệu Revit sẽ được lấy khi bấm Refresh Elements hoặc Refresh Dropdown Values.")

    # -------------------------------------------------------------------------
    # Init
    # -------------------------------------------------------------------------

    def init_options(self):
        reset_observable(self.OperatorOptions, OPERATOR_OPTIONS)
        reset_observable(self.ValueSourceOptions, VALUE_SOURCE_OPTIONS)
        reset_observable(self.ValueTypeOptions, VALUE_TYPE_OPTIONS)
        reset_observable(self.ElementTypeOptions, ELEMENT_TYPE_OPTIONS)
        reset_observable(self.RevitFieldNameOptions, BASE_REVIT_FIELD_OPTIONS)

    def bind_events(self):
        self.btnBrowseMappingCsv.Click += self.on_browse_mapping
        self.btnBrowseRuleCsv.Click += self.on_browse_rules

        self.btnLoadAll.Click += self.on_load_all
        self.btnSaveAll.Click += self.on_save_all
        self.btnSaveAllAs.Click += self.on_save_all_as

        self.btnAddMapping.Click += self.on_add_mapping
        self.btnDeleteMapping.Click += self.on_delete_mapping
        self.btnRenameMapping.Click += self.on_rename_mapping
        self.btnMoveMappingLeft.Click += self.on_move_mapping_left
        self.btnMoveMappingRight.Click += self.on_move_mapping_right
        self.btnRebuildColumns.Click += self.on_rebuild_columns
        self.btnValidateMapping.Click += self.on_validate_mapping

        self.btnAddRule.Click += self.on_add_rule
        self.btnDeleteRule.Click += self.on_delete_rule
        self.btnDuplicateRule.Click += self.on_duplicate_rule
        self.btnRefreshRuleGrid.Click += self.on_refresh_rule_grid

        self.btnRefreshDropdowns.Click += self.on_refresh_dropdowns
        self.btnRefreshElements.Click += self.on_refresh_elements
        self.btnApplyRules.Click += self.on_apply_rules

        self.btnValidate.Click += self.on_validate_all
        self.btnClose.Click += self.on_close

        # Không dùng SelectionChanged / CurrentCellChanged để rebuild cột liên tục.
        # Hai event đó bắn rất nhiều khi user click trong grid, dễ làm UI Not Responding.
        # Chỉ bắt event kết thúc chỉnh sửa và LostFocus, sau đó debounce rebuild bằng Dispatcher.
        self.dgMapping.CellEditEnding += self.on_mapping_edit_ending
        self.dgMapping.RowEditEnding += self.on_mapping_edit_ending
        self.dgMapping.LostFocus += self.on_mapping_edit_ending

        self.dgRules.CellEditEnding += self.on_rule_edit_ending
        self.dgRules.RowEditEnding += self.on_rule_edit_ending

        self.txtRuleSearch.TextChanged += self.on_rule_search_changed
        self.btnClearRuleSearch.Click += self.on_clear_rule_search

        self.txtElementSearch.TextChanged += self.on_element_search_changed
        self.cmbElementFilter.SelectionChanged += self.on_element_filter_changed

        self.Closed += self.on_closed

    def load_default_files_if_available(self):
        if os.path.exists(self.mapping_csv_path):
            self.load_mapping_csv(self.mapping_csv_path)
        else:
            self.create_default_mapping()

        if os.path.exists(self.rule_csv_path):
            self.load_rule_csv(self.rule_csv_path)
        else:
            self.create_default_rules()

    # -------------------------------------------------------------------------
    # Status / utility
    # -------------------------------------------------------------------------

    def set_status(self, message):
        try:
            self.txtStatus.Text = message
        except:
            pass

    def commit_grids(self):
        try:
            self.dgMapping.CommitEdit()
        except:
            pass

        try:
            self.dgRules.CommitEdit()
        except:
            pass

    def raise_external(self, request):
        self.handler.request = request
        self.external_event.Raise()

    def get_element_filter(self):
        try:
            return clean_text(self.cmbElementFilter.SelectedItem.Content)
        except:
            return "Both"

    def selected_mapping(self):
        try:
            return self.dgMapping.SelectedItem
        except:
            return None

    def get_apply_scope(self):
        try:
            return clean_text(self.cmbApplyScope.SelectedItem.Content)
        except:
            return "Active View"

    # -------------------------------------------------------------------------
    # CSV load / save
    # -------------------------------------------------------------------------

    def mapping_headers(self):
        return MAPPING_FIELDS

    def rule_headers(self):
        headers = []
        headers.append("Enabled")
        headers.append("RuleId")
        headers.append("ElementType")

        for mapping in self.Mappings:
            col = clean_text(mapping.RuleColumnName)
            if col and col not in headers:
                headers.append(col)

        headers.extend([
            "ThicknessMM",
            "InsulationTypeName",
            "Priority",
            "Note"
        ])

        return headers

    def create_default_mapping(self):
        self.Mappings.Clear()

        defaults = [
            ("TRUE", "System", "System Name", "Equals", "RevitDropdown", "Text", ""),
            ("TRUE", "Level", "Level Name", "Equals", "RevitDropdown", "Text", ""),
            ("TRUE", "Size", "Overall Size", "Equals", "RevitDropdown", "SizeText", ""),
            ("TRUE", "Area", "Installation Area", "Contains", "ManualOrDropdown", "Text", ""),
            ("TRUE", "Workset", "Workset", "Equals", "RevitDropdown", "Text", "")
        ]

        for item in defaults:
            row = {
                "Enabled": item[0],
                "RuleColumnName": item[1],
                "RevitFieldName": item[2],
                "Operator": item[3],
                "ValueSource": item[4],
                "ValueType": item[5],
                "Note": item[6]
            }
            self.Mappings.Add(mapping_from_dict(row))

        self.dgMapping.ItemsSource = self.Mappings

    def create_default_rules(self):
        self.Rules.Clear()

        rule = RuleItem()
        rule.Enabled = True
        rule.RuleId = "R001"
        rule.ElementType = "Both"
        rule.ThicknessMM = "30"
        rule.InsulationTypeName = ""
        rule.Priority = "100"
        rule.Note = "Default rule"

        for mapping in self.Mappings:
            rule.set_dynamic_value(mapping, "")

        self.Rules.Add(rule)
        self.refresh_filtered_rules()

    def load_mapping_csv(self, path):
        rows = read_csv_dicts(path)

        old_suppress = self._suppress_mapping_events
        self._suppress_mapping_events = True

        self.Mappings.Clear()

        for row in rows:
            self.Mappings.Add(mapping_from_dict(row))

        self.mapping_csv_path = path
        self.txtMappingCsvPath.Text = path
        self.dgMapping.ItemsSource = self.Mappings
        if not getattr(self, "_is_loading_ui", False):
            self.rebuild_rule_grid_columns()

        self._suppress_mapping_events = old_suppress
        self.set_status("Loaded mapping CSV: {0}".format(path))

    def save_mapping_csv(self, path):
        self.commit_grids()

        rows = []
        for mapping in self.Mappings:
            mapping.update_binding_name()
            rows.append(mapping.to_dict())

        write_csv_dicts(path, self.mapping_headers(), rows)

        self.mapping_csv_path = path
        self.txtMappingCsvPath.Text = path

        self.set_status("Saved mapping CSV: {0}".format(path))

    def load_rule_csv(self, path):
        rows = read_csv_dicts(path)

        old_suppress = self._suppress_mapping_events
        self._suppress_mapping_events = True
        headers = read_csv_headers(path)

        # Nếu rule CSV có cột động chưa có mapping, tự tạo mapping tạm.
        existing = {}
        for m in self.Mappings:
            existing[clean_text(m.RuleColumnName).lower()] = True

        for header in headers:
            h = clean_text(header)
            if h.lower() in FIXED_RULE_FIELD_SET:
                continue
            if h.lower() in existing:
                continue

            m = MappingItem()
            m.Enabled = True
            m.RuleColumnName = h
            m.RevitFieldName = h
            m.Operator = "Equals"
            m.ValueSource = "ManualOrDropdown"
            m.ValueType = "Auto"
            m.update_binding_name()
            self.Mappings.Add(m)
            existing[h.lower()] = True

        self.Rules.Clear()

        for row in rows:
            self.Rules.Add(rule_from_dict(row, self.Mappings))

        self.rule_csv_path = path
        self.txtRuleCsvPath.Text = path
        self.refresh_filtered_rules()

        # Giữ dropdown values đã lưu trong CSV để load lại vẫn hiển thị đúng,
        # kể cả khi chưa refresh Revit hoặc active view hiện tại không có value đó.
        self.add_current_rule_values_to_dropdown_data()

        if not getattr(self, "_is_loading_ui", False):
            self.rebuild_rule_grid_columns()

        self._suppress_mapping_events = old_suppress
        self.set_status("Loaded rule CSV: {0}".format(path))

    def save_rule_csv(self, path):
        self.commit_grids()

        headers = self.rule_headers()
        rows = []

        for rule in self.Rules:
            rows.append(rule.to_dict(self.Mappings))

        write_csv_dicts(path, headers, rows)

        self.rule_csv_path = path
        self.txtRuleCsvPath.Text = path

        self.set_status("Saved rule CSV: {0}".format(path))

    # -------------------------------------------------------------------------
    # Dynamic grid column builders
    # -------------------------------------------------------------------------

    def make_combo_edit_style(self):
        try:
            from System.Windows import Style
            style = Style(ComboBox)
            style.Setters.Add(Setter(ComboBox.IsEditableProperty, True))
            style.Setters.Add(Setter(ComboBox.IsTextSearchEnabledProperty, True))
            return style
        except:
            return None

    def add_text_column(self, grid, header, binding_name, width):
        col = DataGridTextColumn()
        col.Header = header
        col.Binding = WpfBinding(binding_name)
        col.Width = dg_width(width)
        grid.Columns.Add(col)

    def add_checkbox_column(self, grid, header, binding_name, width):
        col = DataGridCheckBoxColumn()
        col.Header = header
        col.Binding = WpfBinding(binding_name)
        col.Width = dg_width(width)
        grid.Columns.Add(col)

    def add_combo_column(self, grid, header, binding_name, items_source, width, editable):
        col = DataGridComboBoxColumn()
        col.Header = header
        col.SelectedItemBinding = WpfBinding(binding_name)
        col.ItemsSource = items_source
        col.Width = dg_width(width)

        if editable:
            style = self.make_combo_edit_style()
            if style is not None:
                col.EditingElementStyle = style

        grid.Columns.Add(col)

    def make_observable(self, values):
        col = ObservableCollection[object]()
        for value in values:
            col.Add(value)
        return col

    def row_options_binding_name(self, mapping):
        mapping.update_binding_name()
        return mapping.BindingName + "_Options"

    def make_two_way_binding(self, binding_name, update_property_changed=True):
        b = WpfBinding(binding_name)
        try:
            b.Mode = BindingMode.TwoWay
        except:
            pass
        if update_property_changed:
            try:
                b.UpdateSourceTrigger = UpdateSourceTrigger.PropertyChanged
            except:
                pass
        return b

    def add_row_combo_column(self, grid, header, binding_name, items_source_binding_name, width, editable):
        """
        Tạo ComboBox column có ItemsSource theo từng row.

        Đây là điểm sửa V4.0:
        - DataGridComboBoxColumn cũ dùng chung 1 list cho cả cột.
        - Bây giờ mỗi RuleItem có list riêng, phụ thuộc ElementType của rule.
        - Vì vậy Pipe Fitting + Family Name chỉ xổ Family Name của Pipe Fitting.
        """
        col = DataGridTextColumn()

        try:
            from System.Windows.Controls import DataGridTemplateColumn
            col = DataGridTemplateColumn()
            col.Header = header
            col.Width = dg_width(width)

            template = DataTemplate()
            factory = FrameworkElementFactory(ComboBox)
            factory.SetValue(ComboBox.IsEditableProperty, bool(editable))
            factory.SetValue(ComboBox.IsTextSearchEnabledProperty, True)
            factory.SetBinding(ComboBox.ItemsSourceProperty, WpfBinding(items_source_binding_name))

            if editable:
                factory.SetBinding(ComboBox.TextProperty, self.make_two_way_binding(binding_name))
            else:
                factory.SetBinding(ComboBox.SelectedItemProperty, self.make_two_way_binding(binding_name))

            template.VisualTree = factory
            col.CellTemplate = template
            grid.Columns.Add(col)
            return
        except:
            # Fallback an toàn nếu runtime WPF không cho tạo DataTemplate bằng code.
            self.add_text_column(grid, header, binding_name, width)

    def update_rule_dropdown_options(self, rule):
        """
        Cập nhật dropdown riêng cho từng dòng rule theo ElementType của dòng đó.
        Không thay đổi dữ liệu rule, chỉ cập nhật datasource cho UI.
        """
        if rule is None:
            return

        elem_type = clean_text(rule.ElementType)

        for mapping in self.Mappings:
            mapping.update_binding_name()
            values = self.get_dropdown_for_field(mapping.RevitFieldName, elem_type)

            current_value = rule.get_dynamic_value(mapping)
            if clean_text(current_value) and clean_text(current_value) not in [clean_text(v) for v in values]:
                values.append(current_value)

            setattr(rule, self.row_options_binding_name(mapping), self.make_observable(values))

        # V4.3: InsulationTypeName luôn lấy dropdown theo ElementType của từng rule.
        # Pipe/Pipe Fitting/Pipe Accessory/Flex Pipe -> Pipe insulation types.
        # Duct/Duct Fitting/Duct Accessory/Flex Duct -> Duct insulation types.
        # Không dùng global list ở đây để tránh lẫn Pipe Insulation và Duct Insulation.
        ins_values = self.get_dropdown_for_field("InsulationTypeName", elem_type)
        if clean_text(rule.InsulationTypeName) and clean_text(rule.InsulationTypeName) not in [clean_text(v) for v in ins_values]:
            ins_values.append(rule.InsulationTypeName)
        setattr(rule, "InsulationTypeName_Options", self.make_observable(ins_values))

    def update_all_rule_dropdown_options(self):
        for rule in self.Rules:
            self.update_rule_dropdown_options(rule)

    def get_rule_search_text(self):
        try:
            return clean_text(self.txtRuleSearch.Text).lower()
        except:
            return ""

    def rule_matches_search(self, rule, search):
        if search == "":
            return True

        values = [
            bool_to_csv(rule.Enabled),
            rule.RuleId,
            rule.ElementType,
            rule.ThicknessMM,
            rule.InsulationTypeName,
            rule.Priority,
            rule.Note
        ]

        for mapping in self.Mappings:
            values.append(mapping.RuleColumnName)
            values.append(rule.get_dynamic_value(mapping))

        blob = " ".join([clean_text(v) for v in values]).lower()
        return search in blob

    def refresh_filtered_rules(self):
        search = self.get_rule_search_text()
        self.FilteredRules.Clear()

        self.update_all_rule_dropdown_options()

        for rule in self.Rules:
            if self.rule_matches_search(rule, search):
                self.FilteredRules.Add(rule)

        try:
            self.dgRules.ItemsSource = self.FilteredRules
            self.dgRules.Items.Refresh()
        except:
            pass

    def add_current_rule_values_to_dropdown_data(self):
        """
        Bổ sung giá trị đang có trong insulation_rules.csv vào dropdown hiện tại.

        V4.0:
        - Thêm vào global dropdown.
        - Thêm vào bucket theo ElementType của từng rule.
        Nhờ vậy Load Rules vẫn giữ đúng dropdown theo từng ElementType, kể cả chưa refresh Revit.
        """
        try:
            for rule in self.Rules:
                elem_type = clean_text(rule.ElementType)
                type_bucket = ensure_dropdown_type_bucket(self.dropdown_data, elem_type)

                for mapping in self.Mappings:
                    field = clean_text(mapping.RevitFieldName)
                    if field == "":
                        continue

                    value = rule.get_dynamic_value(mapping)
                    if clean_text(value) == "":
                        continue

                    add_dropdown_value(self.dropdown_data, field, value)
                    add_dropdown_value(type_bucket, field, value)

                if clean_text(rule.InsulationTypeName):
                    add_dropdown_value(self.dropdown_data, "InsulationTypeName", rule.InsulationTypeName)
                    add_dropdown_value(type_bucket, "InsulationTypeName", rule.InsulationTypeName)
        except:
            pass


    def get_dropdown_for_field(self, field_name, element_type=None):
        """
        Lấy dropdown values cho field.

        Nếu element_type có giá trị, ưu tiên lấy bucket theo ElementType.
        Ví dụ:
        - element_type = Pipe Fitting
        - field_name = Family Name
        -> chỉ trả về Family Name của Pipe Fitting đã collect.

        Nếu không có bucket theo ElementType, trả về list rỗng để người dùng vẫn nhập tay
        nếu ValueSource = ManualOrDropdown. Không fallback sang global để tránh lẫn dữ liệu sai loại.
        """
        field = clean_text(field_name)
        elem_type = clean_text(element_type)

        if field == "":
            return []

        if elem_type != "":
            values = {}
            by_type = self.dropdown_data.get(DROPDOWN_BY_ELEMENT_TYPE_KEY, {})

            for key in matching_element_type_keys(elem_type):
                bucket = by_type.get(key, None)
                if bucket and field in bucket:
                    for value in bucket[field].keys():
                        add_unique_text(values, value)

            return sorted_keys(values)

        if field in self.dropdown_data:
            return sorted_keys(self.dropdown_data[field])

        return []


    def rebuild_rule_grid_columns(self):
        """
        Dựng lại các cột của Rule DataGrid từ bảng Parameter Column Mapping.

        Lưu ý sửa lỗi treo UI:
        - Không gọi commit_grids() bên trong hàm này.
          CommitEdit trong lúc DataGrid đang phát event có thể tạo vòng lặp event.
        - Luôn dùng try/finally để trả cờ _is_rebuilding_columns về False.
          Nếu lỗi xảy ra giữa chừng mà cờ không trả về, UI có thể bị trạng thái treo giả.
        """
        if self._is_rebuilding_columns:
            return

        self._is_rebuilding_columns = True

        try:
            self.dgRules.Columns.Clear()

            self.add_checkbox_column(self.dgRules, "Enabled", "Enabled", 72)
            self.add_text_column(self.dgRules, "RuleId", "RuleId", 85)
            self.add_combo_column(self.dgRules, "ElementType", "ElementType", self.ElementTypeOptions, 105, False)

            for mapping in self.Mappings:
                mapping.update_binding_name()

                if not SHOW_DISABLED_MAPPING_COLUMNS and not mapping.Enabled:
                    continue

                col_name = clean_text(mapping.RuleColumnName)
                if col_name == "":
                    continue

                header = col_name
                if not mapping.Enabled:
                    header = col_name + " (off)"

                field = clean_text(mapping.RevitFieldName)
                value_source = clean_text(mapping.ValueSource)
                if value_source in ["RevitDropdown", "ManualOrDropdown"]:
                    self.add_row_combo_column(
                        self.dgRules,
                        header,
                        mapping.BindingName,
                        self.row_options_binding_name(mapping),
                        175,
                        value_source == "ManualOrDropdown"
                    )
                else:
                    self.add_text_column(self.dgRules, header, mapping.BindingName, 175)

            self.add_text_column(self.dgRules, "ThicknessMM", "ThicknessMM", 105)

            self.add_row_combo_column(
                self.dgRules,
                "InsulationTypeName",
                "InsulationTypeName",
                "InsulationTypeName_Options",
                190,
                True
            )

            self.add_text_column(self.dgRules, "Priority", "Priority", 80)
            self.add_text_column(self.dgRules, "Note", "Note", 260)

            self.refresh_filtered_rules()

        finally:
            self._is_rebuilding_columns = False

    def active_preview_fixed_fields(self):
        return [
            ("ElementId", "ElementId", 85),
            ("Category", "Category", 120),
            ("System Name", "SystemName", 180),
            ("System Abbreviation", "SystemAbbreviation", 140),
            ("System Classification", "SystemClassification", 175),
            ("Level Name", "LevelName", 130),
            ("Workset", "Workset", 140),
            ("Phase", "Phase", 120),
            ("Family Name", "FamilyName", 150),
            ("Type Name", "TypeName", 170),
            ("Diameter", "Diameter", 90),
            ("Width", "Width", 90),
            ("Height", "Height", 90),
            ("Length", "Length", 90),
            ("Overall Size", "OverallSize", 130),
            ("Insulation Existing", "InsulationExisting", 140),
            ("Comments", "Comments", 180),
            ("Mark", "Mark", 100)
        ]

    def rebuild_active_element_columns(self, extra_fields):
        self.dgActiveElements.Columns.Clear()

        fixed_items = self.active_preview_fixed_fields()

        for item in fixed_items:
            self.add_text_column(self.dgActiveElements, item[0], item[1], item[2])

        already = {}
        for item in fixed_items:
            already[item[0]] = True

        for field in extra_fields:
            field = clean_text(field)
            if field == "":
                continue
            if field in already:
                continue

            self.add_text_column(self.dgActiveElements, field, safe_binding_name(field), 160)
            already[field] = True

    # -------------------------------------------------------------------------
    # Mapping actions
    # -------------------------------------------------------------------------

    def on_mapping_edit_ending(self, sender, args):
        """
        Event nhẹ cho bảng Parameter Column Mapping.

        Lỗi cũ:
        - Mỗi click / selection trong mapping grid đều gọi rebuild Rule DataGrid.
        - rebuild lại gọi CommitEdit và Items.Refresh, làm WPF phát event tiếp.
        - Kết quả: Not Responding, loading liên tục, thậm chí văng Revit.

        Cách mới:
        - Chỉ đánh dấu mapping đã thay đổi.
        - Đưa rebuild vào Dispatcher.Background để chạy sau khi WPF commit xong cell.
        - Nếu nhiều event bắn liên tiếp, chỉ rebuild một lần.
        """
        self.request_mapping_rebuild("Mapping edited")

    def on_mapping_changed(self, sender, args):
        """
        Giữ lại hàm cũ để tương thích nếu chỗ khác còn gọi.
        """
        self.request_mapping_rebuild("Mapping changed")

    def request_mapping_rebuild(self, reason):
        if getattr(self, "_is_loading_ui", False):
            return

        if getattr(self, "_suppress_mapping_events", False):
            return

        if getattr(self, "_is_rebuilding_columns", False):
            return

        self._mapping_dirty = True

        if getattr(self, "_mapping_rebuild_pending", False):
            return

        self._mapping_rebuild_pending = True

        try:
            self.Dispatcher.BeginInvoke(
                DispatcherPriority.Background,
                Action(self.apply_pending_mapping_rebuild)
            )
        except:
            self.apply_pending_mapping_rebuild()

    def apply_pending_mapping_rebuild(self):
        self._mapping_rebuild_pending = False

        if not getattr(self, "_mapping_dirty", False):
            return

        self._mapping_dirty = False

        if getattr(self, "_is_loading_ui", False):
            return

        if getattr(self, "_is_rebuilding_columns", False):
            return

        self._suppress_mapping_events = True

        try:
            try:
                self.commit_grids()
            except:
                pass

            for m in self.Mappings:
                m.update_binding_name()

            self.rebuild_rule_grid_columns()
            self.set_status("Mapping updated. Rule columns refreshed once.")

        finally:
            self._suppress_mapping_events = False

    def on_rebuild_columns(self, sender, args):
        self._suppress_mapping_events = True
        try:
            self.commit_grids()

            for m in self.Mappings:
                m.update_binding_name()

            self.rebuild_rule_grid_columns()
            self.set_status("Rule columns rebuilt from mapping.")
        finally:
            self._suppress_mapping_events = False

    def on_add_mapping(self, sender, args):
        name = forms.ask_for_string(
            default="NewColumn",
            prompt="Nhap RuleColumnName cho cot moi",
            title="Add Parameter Column"
        )

        if name is None:
            return

        name = clean_text(name)
        if name == "":
            alert("RuleColumnName cannot be empty.")
            return

        m = MappingItem()
        m.Enabled = True
        m.RuleColumnName = name
        m.RevitFieldName = ""
        m.Operator = "Equals"
        m.ValueSource = "ManualOrDropdown"
        m.ValueType = "Auto"
        m.update_binding_name()

        self.Mappings.Add(m)
        self.dgMapping.SelectedItem = m

        for rule in self.Rules:
            rule.set_dynamic_value(m, "")

        self.rebuild_rule_grid_columns()
        try:
            self.dgMapping.Items.Refresh()
        except:
            pass

        self.set_status("Added parameter column: {0}".format(name))

    def on_delete_mapping(self, sender, args):
        selected = list(self.dgMapping.SelectedItems)

        if len(selected) == 0:
            alert("Select mapping row to delete.")
            return

        ok = yes_no("Delete selected mapping column(s)? Rule CSV data in these dynamic columns will be removed on next save.")
        if not ok:
            return

        for m in selected:
            try:
                self.Mappings.Remove(m)
            except:
                pass

        self.rebuild_rule_grid_columns()
        self.set_status("Deleted {0} mapping column(s).".format(len(selected)))

    def on_rename_mapping(self, sender, args):
        m = self.selected_mapping()

        if m is None:
            alert("Select mapping row to rename.")
            return

        old_name = clean_text(m.RuleColumnName)

        new_name = forms.ask_for_string(
            default=old_name,
            prompt="New RuleColumnName",
            title="Rename Rule Column"
        )

        if new_name is None:
            return

        new_name = clean_text(new_name)

        if new_name == "":
            alert("RuleColumnName cannot be empty.")
            return

        old_binding = m.BindingName
        m.RuleColumnName = new_name
        m.update_binding_name()
        new_binding = m.BindingName

        for rule in self.Rules:
            try:
                old_value = getattr(rule, old_binding)
            except:
                old_value = ""
            setattr(rule, new_binding, old_value)

        self.rebuild_rule_grid_columns()
        try:
            self.dgMapping.Items.Refresh()
            self.dgRules.Items.Refresh()
        except:
            pass
        self.set_status("Renamed column {0} to {1}.".format(old_name, new_name))

    def move_mapping(self, direction):
        m = self.selected_mapping()

        if m is None:
            alert("Select mapping row to move.")
            return

        index = self.Mappings.IndexOf(m)
        new_index = index + direction

        if new_index < 0 or new_index >= self.Mappings.Count:
            return

        self.Mappings.RemoveAt(index)
        self.Mappings.Insert(new_index, m)
        self.dgMapping.SelectedItem = m
        self.rebuild_rule_grid_columns()

    def on_move_mapping_left(self, sender, args):
        self.move_mapping(-1)

    def on_move_mapping_right(self, sender, args):
        self.move_mapping(1)

    def on_validate_mapping(self, sender, args):
        self.commit_grids()
        errors = validate_mapping(self.Mappings)

        if len(errors) == 0:
            alert("Mapping validate OK.")
            self.set_status("Mapping validate OK.")
            return

        output.print_md("# Mapping Validation Report")
        for err in errors:
            print("- {0}".format(err))

        alert("Mapping validation found {0} error(s). See pyRevit output.".format(len(errors)))

    # -------------------------------------------------------------------------
    # Rule actions
    # -------------------------------------------------------------------------

    def next_rule_id(self):
        max_num = 0

        for rule in self.Rules:
            rid = clean_text(rule.RuleId).upper()
            if rid.startswith("R"):
                num = try_float(rid[1:])
                if num is not None:
                    max_num = max(max_num, int(num))

        return "R{0:03d}".format(max_num + 1)

    def on_add_rule(self, sender, args):
        rule = RuleItem()
        rule.Enabled = True
        rule.RuleId = self.next_rule_id()
        rule.ElementType = "Both"
        rule.ThicknessMM = "30"
        rule.Priority = "100"

        for m in self.Mappings:
            rule.set_dynamic_value(m, "")

        self.Rules.Add(rule)
        self.refresh_filtered_rules()
        self.dgRules.SelectedItem = rule

        self.set_status("Added rule {0}.".format(rule.RuleId))

    def on_delete_rule(self, sender, args):
        selected = list(self.dgRules.SelectedItems)

        if len(selected) == 0:
            alert("Select rule to delete.")
            return

        for rule in selected:
            try:
                self.Rules.Remove(rule)
            except:
                pass

        self.refresh_filtered_rules()
        self.set_status("Deleted {0} rule(s).".format(len(selected)))

    def on_duplicate_rule(self, sender, args):
        try:
            source = self.dgRules.SelectedItem
        except:
            source = None

        if source is None:
            alert("Select rule to duplicate.")
            return

        rule = RuleItem()
        rule.Enabled = source.Enabled
        rule.RuleId = self.next_rule_id()
        rule.ElementType = source.ElementType
        rule.ThicknessMM = source.ThicknessMM
        rule.InsulationTypeName = source.InsulationTypeName
        rule.Priority = source.Priority
        rule.Note = source.Note

        for m in self.Mappings:
            rule.set_dynamic_value(m, source.get_dynamic_value(m))

        self.Rules.Add(rule)
        self.refresh_filtered_rules()
        self.dgRules.SelectedItem = rule

        self.set_status("Duplicated rule to {0}.".format(rule.RuleId))

    def on_refresh_rule_grid(self, sender, args):
        self.commit_grids()
        for m in self.Mappings:
            m.update_binding_name()
        self.rebuild_rule_grid_columns()
        try:
            self.dgRules.Items.Refresh()
        except:
            pass
        self.set_status("Rule grid refreshed without reloading CSV.")

    # -------------------------------------------------------------------------
    # CSV events
    # -------------------------------------------------------------------------

    def on_load_all(self, sender, args):
        mapping_path = clean_text(self.txtMappingCsvPath.Text)
        rule_path = clean_text(self.txtRuleCsvPath.Text)

        if mapping_path and os.path.exists(mapping_path):
            self.load_mapping_csv(mapping_path)

        if rule_path and os.path.exists(rule_path):
            self.load_rule_csv(rule_path)

        self.rebuild_rule_grid_columns()
        self.set_status("Loaded mapping and rule CSV.")

    def on_save_all(self, sender, args):
        self.commit_grids()

        mapping_path = clean_text(self.txtMappingCsvPath.Text)
        rule_path = clean_text(self.txtRuleCsvPath.Text)

        if not mapping_path:
            mapping_path = save_csv_file("Save insulation_rule_mapping.csv", "insulation_rule_mapping.csv")
        if not rule_path:
            rule_path = save_csv_file("Save insulation_rules.csv", "insulation_rules.csv")

        if not mapping_path or not rule_path:
            return

        self.save_mapping_csv(mapping_path)
        self.save_rule_csv(rule_path)
        self.set_status("Saved mapping and rule CSV together.")

    def on_save_all_as(self, sender, args):
        self.commit_grids()

        mapping_path = save_csv_file("Save insulation_rule_mapping.csv", "insulation_rule_mapping.csv")
        if not mapping_path:
            return

        rule_path = save_csv_file("Save insulation_rules.csv", "insulation_rules.csv")
        if not rule_path:
            return

        self.save_mapping_csv(mapping_path)
        self.save_rule_csv(rule_path)
        self.set_status("Saved mapping and rule CSV as new files.")

    def on_browse_mapping(self, sender, args):
        path = pick_csv_file("Select insulation_rule_mapping.csv")
        if path:
            self.mapping_csv_path = path
            self.txtMappingCsvPath.Text = path

    def on_browse_rules(self, sender, args):
        path = pick_csv_file("Select insulation_rules.csv")
        if path:
            self.rule_csv_path = path
            self.txtRuleCsvPath.Text = path

    def on_load_mapping(self, sender, args):
        path = clean_text(self.txtMappingCsvPath.Text)

        if not path:
            path = pick_csv_file("Select insulation_rule_mapping.csv")

        if not path:
            return

        if not os.path.exists(path):
            alert("Mapping CSV does not exist:\n{0}".format(path))
            return

        self.load_mapping_csv(path)

    def on_save_mapping(self, sender, args):
        self.on_save_all(sender, args)

    def on_load_rules(self, sender, args):
        path = clean_text(self.txtRuleCsvPath.Text)

        if not path:
            path = pick_csv_file("Select insulation_rules.csv")

        if not path:
            return

        if not os.path.exists(path):
            alert("Rule CSV does not exist:\n{0}".format(path))
            return

        self.load_rule_csv(path)

    def on_save_rules(self, sender, args):
        self.on_save_all(sender, args)

    def on_save_as_rules(self, sender, args):
        self.on_save_all_as(sender, args)

    # -------------------------------------------------------------------------
    # Validate
    # -------------------------------------------------------------------------

    def print_validation_errors(self, errors):
        output.print_md("# Insulation Rule Validation Report")
        for err in errors:
            print("- {0}".format(err))

    def on_validate_all(self, sender, args):
        self.commit_grids()

        errors = validate_all(self.Mappings, self.Rules)

        if len(errors) == 0:
            alert("Validate OK.")
            self.set_status("Validate OK.")
            return

        self.print_validation_errors(errors)
        alert("Validate found {0} error(s). See pyRevit output.".format(len(errors)))
        self.set_status("Validate failed.")

    # -------------------------------------------------------------------------
    # Rule search / row dropdown refresh
    # -------------------------------------------------------------------------

    def on_rule_search_changed(self, sender, args):
        self.refresh_filtered_rules()

    def on_clear_rule_search(self, sender, args):
        try:
            self.txtRuleSearch.Text = ""
        except:
            pass
        self.refresh_filtered_rules()

    def on_rule_edit_ending(self, sender, args):
        """
        Khi người dùng đổi ElementType trong Rule DataGrid, dropdown của các cột động
        cần đổi theo ElementType mới. Chạy trễ bằng Dispatcher để WPF commit cell trước.
        """
        if getattr(self, "_rule_options_refresh_pending", False):
            return

        self._rule_options_refresh_pending = True

        try:
            self.Dispatcher.BeginInvoke(
                DispatcherPriority.Background,
                Action(self.apply_pending_rule_options_refresh)
            )
        except:
            self.apply_pending_rule_options_refresh()

    def apply_pending_rule_options_refresh(self):
        self._rule_options_refresh_pending = False

        try:
            self.commit_grids()
        except:
            pass

        self.refresh_filtered_rules()
        self.set_status("Rule dropdown values refreshed by ElementType.")

    # -------------------------------------------------------------------------
    # Active View Elements
    # -------------------------------------------------------------------------

    def get_search_text(self):
        try:
            return clean_text(self.txtElementSearch.Text).lower()
        except:
            return ""

    def refresh_filtered_elements(self):
        search = self.get_search_text()
        self.FilteredActiveElements.Clear()

        for item in self.ActiveElements:
            if search == "" or search in item.search_blob():
                self.FilteredActiveElements.Add(item)

        self.dgActiveElements.ItemsSource = self.FilteredActiveElements

    def on_element_search_changed(self, sender, args):
        self.refresh_filtered_elements()

    def on_element_filter_changed(self, sender, args):
        self.on_refresh_elements(sender, args)

    def on_refresh_elements(self, sender, args):
        self.raise_external("REFRESH_ELEMENTS")

    def on_refresh_dropdowns(self, sender, args):
        self.raise_external("REFRESH_DROPDOWNS")

    def update_dropdown_data(self, data):
        self.dropdown_data = data

        # Không làm mất các value đã có trong CSV sau khi refresh Revit.
        # Đây chỉ là datasource cho UI dropdown, không ảnh hưởng logic so sánh.
        self.add_current_rule_values_to_dropdown_data()

        fields = {}
        for item in BASE_REVIT_FIELD_OPTIONS:
            add_unique_text(fields, item)

        for item in sorted_keys(data.get("ParameterNames", {})):
            add_unique_text(fields, item)

        reset_observable(self.RevitFieldNameOptions, sorted_keys(fields))
        self.request_mapping_rebuild("Dropdown data updated")

    def get_extra_preview_fields_from_mapping(self):
        fixed = {}
        for item in [
            "System Name",
            "System Abbreviation",
            "System Classification",
            "Level Name",
            "Workset",
            "Phase",
            "Category",
            "Family Name",
            "Type Name",
            "Diameter",
            "Width",
            "Height",
            "Length",
    "Overall Size",
            "Comments",
            "Mark"
        ]:
            fixed[item] = True

        result = []

        for m in self.Mappings:
            field = clean_text(m.RevitFieldName)
            if field == "":
                continue

            if field in fixed:
                continue

            if field not in result:
                result.append(field)

        return result

    # -------------------------------------------------------------------------
    # ExternalEvent callbacks
    # -------------------------------------------------------------------------

    def external_refresh_dropdowns(self, revit_doc, active_view):
        data = collect_dropdown_data(revit_doc, active_view, self.get_element_filter())
        self.update_dropdown_data(data)
        self.set_status("Dropdown values refreshed by ElementType buckets from active view: {0}".format(get_element_name(active_view)))

    def external_refresh_elements(self, revit_doc, active_view):
        element_filter = self.get_element_filter()
        elems = collect_active_view_pipe_ducts(revit_doc, active_view, element_filter)

        extra_fields = self.get_extra_preview_fields_from_mapping()
        self.rebuild_active_element_columns(extra_fields)

        self.ActiveElements.Clear()

        for elem in elems:
            item = make_preview_item(elem, active_view, extra_fields)
            self.ActiveElements.Add(item)

        data = collect_dropdown_data(revit_doc, active_view, element_filter)
        self.update_dropdown_data(data)

        self.refresh_filtered_elements()

        self.txtActiveViewInfo.Text = "Active View: {0} | Elements: {1}".format(
            get_element_name(active_view),
            len(elems)
        )

        self.set_status("Refreshed {0} element(s) from active view.".format(len(elems)))

    def external_apply_rules(self, revit_doc, active_view, active_uidoc):
        errors = validate_all(self.Mappings, self.Rules)
        if len(errors) > 0:
            self.print_validation_errors(errors)
            alert("Validation failed. See pyRevit output.")
            return

        scope = self.get_apply_scope()
        # V4.3: Apply Rules collect theo Rule.ElementType, khong theo Element Filter preview.
        # Nhờ vậy Pipe Fitting / Accessory vẫn được kiểm tra dù Element Filter đang là Pipe/Both.
        elems = collect_apply_scope_elements(revit_doc, active_view, active_uidoc, scope, self.Rules)

        created = 0
        updated = 0
        skipped = 0
        messages = []

        t = Transaction(revit_doc, "Apply Insulation Rules")

        try:
            t.Start()

            for elem in elems:
                report = []
                rule = find_matching_rule(elem, self.Rules, self.Mappings, active_view, report)

                if rule is None:
                    skipped += 1
                    continue

                ok, msg, action = create_insulation_for_element(revit_doc, elem, rule)

                if action == "created":
                    created += 1
                elif action == "updated":
                    updated += 1
                else:
                    skipped += 1

                if msg:
                    messages.append("Element {0}, Rule {1}, Action {2}: {3}".format(
                        element_id_value(elem.Id),
                        rule.RuleId,
                        action,
                        msg
                    ))

            if created + updated > 0:
                t.Commit()
            else:
                t.RollBack()

        except Exception as ex:
            try:
                t.RollBack()
            except:
                pass

            messages.append(str(ex))
            messages.append(traceback.format_exc())

        output.print_md("# Apply Insulation Rules Report")
        print("Active View: {0}".format(get_element_name(active_view)))
        print("Apply Scope: {0}".format(scope))
        print("Element collection: from enabled rule ElementType, not preview Element Filter")
        print("Elements checked: {0}".format(len(elems)))
        print("Insulation created: {0}".format(created))
        print("Insulation updated: {0}".format(updated))
        print("Skipped: {0}".format(skipped))

        if messages:
            output.print_md("## Messages")
            for msg in messages:
                print("- {0}".format(msg))

        alert("Apply Rules done.\nCreated: {0}\nUpdated: {1}\nSkipped: {2}\nSee pyRevit output.".format(created, updated, skipped))

    # -------------------------------------------------------------------------
    # Button events
    # -------------------------------------------------------------------------

    def on_apply_rules(self, sender, args):
        self.commit_grids()
        self.raise_external("APPLY_RULES")

    def on_close(self, sender, args):
        self.Close()

    def on_closed(self, sender, args):
        global _RULE_EDITOR_WINDOW
        global _RULE_EDITOR_HANDLER
        global _RULE_EDITOR_EVENT

        try:
            self.external_event.Dispose()
        except:
            pass

        _RULE_EDITOR_WINDOW = None
        _RULE_EDITOR_HANDLER = None
        _RULE_EDITOR_EVENT = None


# =============================================================================
# MAIN
# =============================================================================

def main():
    global _RULE_EDITOR_WINDOW
    global _RULE_EDITOR_HANDLER
    global _RULE_EDITOR_EVENT

    try:
        if _RULE_EDITOR_WINDOW is not None:
            try:
                _RULE_EDITOR_WINDOW.Activate()
                return
            except:
                pass

        _RULE_EDITOR_HANDLER = RuleEditorExternalEventHandler()
        _RULE_EDITOR_EVENT = ExternalEvent.Create(_RULE_EDITOR_HANDLER)

        _RULE_EDITOR_WINDOW = RuleEditorWindow(_RULE_EDITOR_HANDLER, _RULE_EDITOR_EVENT)
        _RULE_EDITOR_HANDLER.window = _RULE_EDITOR_WINDOW

        # Modeless window: không khóa Revit.
        _RULE_EDITOR_WINDOW.Show()

        # Giữ cấu trúc V3.5 nhưng không tự quét Revit khi mở tool.
        # Đây là nguyên nhân chính làm startup chậm trên project lớn.
        if AUTO_REFRESH_DROPDOWNS_ON_STARTUP:
            _RULE_EDITOR_HANDLER.request = "REFRESH_DROPDOWNS"
            _RULE_EDITOR_EVENT.Raise()
        else:
            try:
                _RULE_EDITOR_WINDOW.set_status("Ready. Bấm Refresh Elements / Refresh Dropdown Values khi cần lấy dữ liệu Revit.")
            except:
                pass

    except Exception as ex:
        output.print_md("# Rule Editor Startup Error")
        print(str(ex))
        print(traceback.format_exc())
        alert("Rule Editor failed. See pyRevit output.")


main()
