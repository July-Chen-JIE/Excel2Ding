import tkinter as tk
from tkinter import ttk


def _bootstrap():
    try:
        import ttkbootstrap as tb
        from ttkbootstrap.widgets import DateEntry as TBDateEntry
        return tb, TBDateEntry
    except Exception:
        try:
            import sys
            import subprocess
            subprocess.run([sys.executable, "-m", "pip", "install", "ttkbootstrap"], check=True)
            import ttkbootstrap as tb
            from ttkbootstrap.widgets import DateEntry as TBDateEntry
            return tb, TBDateEntry
        except Exception:
            return None, None


_tb, _TBDateEntry = _bootstrap()

_BUTTON_STYLE_MAP = {
    'primary': 'Primary.TButton',
    'danger': 'Danger.TButton',
    'info': 'Info.TButton'
}


def make_button(parent, text, command, width, role='primary'):
    if _tb is not None:
        return _tb.Button(parent, text=text, command=command, width=width, bootstyle=role)
    style_name = _BUTTON_STYLE_MAP.get(role, 'TButton')
    return ttk.Button(parent, text=text, command=command, width=width, style=style_name)


def make_date_entry(parent, **kwargs):
    if _TBDateEntry is not None:
        return _TBDateEntry(parent, **kwargs)
    from tkcalendar import DateEntry
    dateformat = kwargs.pop('dateformat', None)
    if dateformat:
        kwargs['date_pattern'] = dateformat.replace('%Y', 'yyyy').replace('%m', 'MM').replace('%d', 'dd')
    kwargs.pop('bootstyle', None)
    return DateEntry(parent, **kwargs)


def set_date_value(widget, value):
    for attempt in (lambda: widget.set_date(value),
                    lambda: _set_entry_value(widget.entry, value),
                    lambda: _set_entry_value(widget, value)):
        try:
            attempt()
            return
        except Exception:
            continue


def _set_entry_value(widget, value):
    widget.delete(0, tk.END)
    widget.insert(0, value)
