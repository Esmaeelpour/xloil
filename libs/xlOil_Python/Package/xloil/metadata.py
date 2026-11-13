"""
    Workbook metadata: read and write the document properties Excel stores in a
    workbook, and a JSON key/value store saved invisibly inside the workbook file.

    ::

        import xloil as xlo

        meta = xlo.workbook_metadata()          # active workbook
        meta.properties['Title'] = "Q3 Report"  # built-in document properties
        meta.custom['Model version'] = "2.1"    # custom document properties
        meta.data['last_run'] = {"user": "me", "rows": 1023}  # JSON store
        print(meta.data.get('last_run'))

    The JSON store and the custom properties are saved with the workbook (so the
    workbook must be saved afterwards for changes to persist). The JSON store is
    held in a Custom XML Part, which Excel keeps intact but does not display.

    All functions here use the COM API, so they must be called from a macro,
    command, event handler or a macro-type worksheet function (the default for
    `xloil.func`) - not during calculation of a non-macro function.
"""

import base64
import json
from collections.abc import MutableMapping
from typing import Any, Iterator, Optional
from xml.sax.saxutils import quoteattr
import xml.etree.ElementTree as ET

import xloil_core

STORE_NAMESPACE = "urn:xloil:metadata:1"

# Office's MsoDocProperties enum
_MSO_PROPERTY_TYPES = {bool: 2, int: 1, float: 5, str: 4}
_MSO_PROPERTY_TYPE_DATE = 3

BUILTIN_PROPERTY_NAMES = (
    "Title", "Subject", "Author", "Keywords", "Comments", "Template",
    "Last Author", "Revision Number", "Application Name", "Last Print Date",
    "Creation Date", "Last Save Time", "Total Editing Time", "Number of Pages",
    "Number of Words", "Number of Characters", "Security", "Category", "Format",
    "Manager", "Company", "Number of Bytes", "Number of Lines", "Number of Paragraphs",
    "Number of Slides", "Number of Notes", "Number of Hidden Slides",
    "Number of Multimedia Clips", "Hyperlink Base", "Number of Characters (with spaces)",
    "Content Type", "Content Status", "Language", "Document Version")


def _workbook_com(workbook):
    """ Resolves None / name / xloil.Workbook / COM Workbook to a COM Workbook """
    if workbook is None:
        return xloil_core.active_workbook().to_com()
    if isinstance(workbook, str):
        return xloil_core.app().workbooks[workbook].to_com()
    to_com = getattr(workbook, "to_com", None)
    if to_com is not None:
        return to_com()
    return workbook


# Note: Office DocumentProperties collections are late-bound. Their `Item` is a
# default member, reachable as props(i) / props(name) but not props.Item(i).

class _BuiltinProperties(MutableMapping):
    """
        Dict-like access to a workbook's built-in document properties such as
        Title, Author, Subject. Some properties are read-only or not available
        until the workbook is saved; these raise KeyError when read.
    """
    def __init__(self, wb_com):
        self._wb = wb_com

    def __getitem__(self, name):
        try:
            return self._wb.BuiltinDocumentProperties(name).Value
        except Exception as e:
            raise KeyError(name) from e

    def __setitem__(self, name, value):
        self._wb.BuiltinDocumentProperties(name).Value = value

    def __delitem__(self, name):
        self[name] = ""

    def __iter__(self) -> Iterator[str]:
        props = self._wb.BuiltinDocumentProperties
        for i in range(1, props.Count + 1):
            yield props(i).Name

    def __len__(self):
        return self._wb.BuiltinDocumentProperties.Count

    def to_dict(self) -> dict:
        """ Returns all properties which currently have a value """
        result = {}
        for name in self:
            try:
                result[name] = self[name]
            except KeyError:
                pass
        return result


class _CustomProperties(MutableMapping):
    """
        Dict-like access to a workbook's custom document properties (visible in
        File > Info > Properties > Advanced Properties > Custom). Values may be
        str, int, float, bool or datetime.
    """
    def __init__(self, wb_com):
        self._wb = wb_com

    def _find(self, name):
        props = self._wb.CustomDocumentProperties
        for i in range(1, props.Count + 1):
            prop = props(i)
            if prop.Name == name:
                return prop
        return None

    def __getitem__(self, name):
        prop = self._find(name)
        if prop is None:
            raise KeyError(name)
        return prop.Value

    def __setitem__(self, name, value):
        import datetime as dt
        prop = self._find(name)
        if prop is not None:
            prop.Delete()
        if isinstance(value, (dt.date, dt.datetime)):
            prop_type = _MSO_PROPERTY_TYPE_DATE
        else:
            # bool before int as bool is a subclass of int
            prop_type = next((t for k, t in _MSO_PROPERTY_TYPES.items()
                              if type(value) is k), None)
            if prop_type is None:
                prop_type = 4
                value = str(value)
        # DocumentProperties is late-bound (no type info), so keyword 
        # arguments are not supported: pass Name, LinkToContent, Type, Value
        self._wb.CustomDocumentProperties.Add(name, False, prop_type, value)

    def __delitem__(self, name):
        prop = self._find(name)
        if prop is None:
            raise KeyError(name)
        prop.Delete()

    def __iter__(self):
        props = self._wb.CustomDocumentProperties
        for i in range(1, props.Count + 1):
            yield props(i).Name

    def __len__(self):
        return self._wb.CustomDocumentProperties.Count


def _read_xml_part(wb_com, namespace: str) -> Optional[str]:
    parts = wb_com.CustomXMLParts.SelectByNamespace(namespace)
    if parts.Count == 0:
        return None
    return parts.Item(1).XML


def _write_xml_part(wb_com, namespace: str, xml: Optional[str]):
    """ Replaces any existing parts in the namespace with `xml` (or deletes if None) """
    parts = wb_com.CustomXMLParts.SelectByNamespace(namespace)
    # Delete in reverse as the collection is live
    for i in range(parts.Count, 0, -1):
        parts.Item(i).Delete()
    if xml is not None:
        wb_com.CustomXMLParts.Add(xml)


def encode_store(values: dict, namespace: str = STORE_NAMESPACE, root: str = "xloilData") -> str:
    """
        Serialises a dict of str -> bytes into the Custom XML format xlOil uses.
        Values are base64 encoded so arbitrary content is safe in XML.
    """
    items = "".join(
        f"<item key={quoteattr(str(k))}>{base64.b64encode(v).decode('ascii')}</item>"
        for k, v in values.items())
    return f'<{root} xmlns="{namespace}">{items}</{root}>'


def decode_store(xml: Optional[str], namespace: str = STORE_NAMESPACE) -> dict:
    """ Inverse of `encode_store` """
    if not xml:
        return {}
    root = ET.fromstring(xml)
    result = {}
    for item in root.findall(f"{{{namespace}}}item"):
        result[item.get("key")] = base64.b64decode(item.text or "")
    return result


class _JsonStore(MutableMapping):
    """
        A dict-like store of JSON-serialisable values, saved inside the workbook
        file in a Custom XML Part. Changes are written to the workbook immediately
        (the workbook must still be saved to keep them).
    """
    def __init__(self, wb_com, namespace=STORE_NAMESPACE):
        self._wb = wb_com
        self._ns = namespace

    def _load(self) -> dict:
        raw = decode_store(_read_xml_part(self._wb, self._ns), self._ns)
        return {k: json.loads(v.decode('utf-8')) for k, v in raw.items()}

    def _save(self, values: dict):
        encoded = {k: json.dumps(v).encode('utf-8') for k, v in values.items()}
        _write_xml_part(self._wb, self._ns, encode_store(encoded, self._ns) if encoded else None)

    def __getitem__(self, key):
        return self._load()[key]

    def __setitem__(self, key, value):
        values = self._load()
        values[key] = value
        self._save(values)

    def __delitem__(self, key):
        values = self._load()
        del values[key]
        self._save(values)

    def __iter__(self):
        return iter(self._load())

    def __len__(self):
        return len(self._load())

    def update(self, other=(), **kwargs):
        # Single write rather than one per key
        values = self._load()
        values.update(other, **kwargs)
        self._save(values)

    def clear(self):
        self._save({})


class WorkbookMetadata:
    """
        Access to a workbook's metadata. Create with `xloil.workbook_metadata`.

        Attributes
        ----------
        properties:
            Built-in document properties (Title, Author, ...) as a dict-like object
        custom:
            Custom document properties as a dict-like object
        data:
            A JSON key/value store saved inside the workbook
    """
    def __init__(self, workbook=None):
        self._wb = _workbook_com(workbook)
        self.properties = _BuiltinProperties(self._wb)
        self.custom = _CustomProperties(self._wb)
        self.data = _JsonStore(self._wb)

    @property
    def name(self) -> str:
        return self._wb.Name

    @property
    def path(self) -> str:
        return self._wb.FullName

    def to_dict(self) -> dict:
        """ Returns all metadata as a plain dict (useful for logging or audit trails) """
        return {
            'name': self.name,
            'path': self.path,
            'properties': self.properties.to_dict(),
            'custom': dict(self.custom),
            'data': dict(self.data),
        }


def workbook_metadata(workbook=None) -> WorkbookMetadata:
    """
        Returns a `WorkbookMetadata` for the given workbook, which may be a name,
        an `xloil.Workbook`, a COM Workbook or None for the active workbook.
    """
    return WorkbookMetadata(workbook)
