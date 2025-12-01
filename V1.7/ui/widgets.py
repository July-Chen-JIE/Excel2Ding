from tkinter import ttk
import tkinter as tk


def _bootstrap():
    try:
        import ttkbootstrap as tb
        from ttkbootstrap.widgets import DateEntry as TBDateEntry
        return tb, TBDateEntry
    except Exception:
        try:
            import sys, subprocess
            subprocess.run([sys.executable, "-m", "pip", "install", "ttkbootstrap"], check=True)
            import ttkbootstrap as tb
            from ttkbootstrap.widgets import DateEntry as TBDateEntry
            return tb, TBDateEntry
        except Exception:
            return None, None


_tb, _TBDateEntry = _bootstrap()


def make_button(parent, text, command, width, role='primary'):
    if _tb is not None:
        return _tb.Button(parent, text=text, command=command, width=width, bootstyle=role)
    style_map = {
        'primary': 'Primary.TButton',
        'danger': 'Danger.TButton',
        'info': 'Info.TButton'
    }
    return ttk.Button(parent, text=text, command=command, width=width, style=style_map.get(role, 'TButton'))


def make_date_entry(parent, **kwargs):
    if _TBDateEntry is not None:
        return _TBDateEntry(parent, **kwargs)
    # 退化到 tkcalendar.DateEntry
    from tkcalendar import DateEntry
    df = kwargs.pop('dateformat', None)
    if df:
        kwargs['date_pattern'] = df.replace('%Y', 'yyyy').replace('%m', 'MM').replace('%d', 'dd')
    kwargs.pop('bootstyle', None)
    return DateEntry(parent, **kwargs)


def set_date_value(widget, value):
    try:
        widget.set_date(value)
        return
    except Exception:
        pass
    try:
        widget.entry.delete(0, tk.END)
        widget.entry.insert(0, value)
        return
    except Exception:
        pass
    try:
        widget.delete(0, tk.END)
        widget.insert(0, value)
    except Exception:
        pass
