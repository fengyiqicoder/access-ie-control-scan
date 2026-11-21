# access-ie-control-scan

Find every legacy (Internet Explorer based) Web Browser Control in a Microsoft Access database, and the VBA lines that will break when Microsoft retires it.

Microsoft is retiring the legacy Web Browser Control in Access in **version 2705 (about 2027-05-18)**, and on **2027-05-11 for Access LTSC 2024**. After that, legacy controls stop loading web content. Microsoft's [announcement](https://techcommunity.microsoft.com/blog/accessblog/deprecation-of-legacy-web-browser-control/4566944) says finding affected controls is currently manual: open each form in Design View, include forms used as subforms. This module does that pass for you.

## Use it (2 minutes)

1. Back up the database.
2. Download [`IEControlScan.bas`](IEControlScan.bas). In Access press `Alt+F11`, then **File > Import File** and pick it.
3. In the Immediate window (`Ctrl+G`) type `ScanIEControls` and press Enter.

You get a summary in the Immediate window and a CSV next to the database (`ie-control-scan-<date>.csv`):

| Type | Object | Control or line | Finding | Action |
|---|---|---|---|---|
| Form | frmCustomer | wbMap | LEGACY Web Browser Control | MUST CONVERT |
| Form | frmMain | sfrmCustomer | Subform control (Hosts: frmCustomer) | INFO |
| Code | Form_frmCustomer | line 42 | Uses .Object of legacy control wbMap | MUST REWRITE |
| Code | modReports | line 118 | MSHTML type | MUST REWRITE |

(Example rows, not from a real database.)

## What it checks

- **Forms and reports**: every control with `ControlType = 128` (`acWebBrowser`, the legacy control). Edge Browser Controls (`134`) are listed as OK. ActiveX `Shell.Explorer` controls are flagged as CHECK: they use the same IE engine but are not the control named in the notice.
- **Subforms and navigation buttons**: listed so you can see which parent forms show an affected subform.
- **VBA** (all modules, form and report modules): `MSHTML`, `HTMLDocument`/`IHTML*` types, `Shell.Explorer`/`InternetExplorer.Application`, `ExecWB`, and any line that uses `.Object`, `.Document` or `Navigate2` on a legacy control found in the scan.

It is read-only. Forms and reports are opened hidden in Design View and closed with `acSaveNo`. Forms you already have open are read in place and left open.

## Rewriting the code

The Edge Browser Control has no `Object` property and no MSHTML document. Common rewrites:

| Legacy control | Edge Browser Control |
|---|---|
| `Me.wb.Object.Navigate2 url` | `Me.wb.Navigate url` (or set `ControlSource`) |
| `Me.wb.Object.Document.getElementById("x").innerText` | `Me.wb.RetrieveJavascriptValue("document.getElementById('x').innerText")` |
| `Me.wb.Object.Document.parentWindow.execScript "f()"` | `Me.wb.ExecuteJavascript "f()"` (runs asynchronously) |
| `Me.wb.Object.Document.body.innerHTML = html` | `Me.wb.ExecuteJavascript "document.body.innerHTML = '...'"` or navigate to a file |
| `Me.wb.Object.Busy` / `ReadyState` | `Me.wb.ReadyState`, or the `DocumentComplete` event |
| `Me.wb.Object.ExecWB` (print, copy) | no direct equivalent; use `ExecuteJavascript "window.print()"` |

Local files and non-HTTPS pages need the Edge control's `TrustedDomains` property set, and paths to local files must be prefixed with `https://msaccess.localhost`. Method reference: [EdgeBrowserControl](https://learn.microsoft.com/office/vba/api/access.edgebrowsercontrol).

## Status

v0.1.0. Written against Microsoft's documented Access object model (`AcControlType`, `EdgeBrowserControl`). It has not yet been run against a wide range of real databases, so please open an issue with what it missed or flagged wrongly. Pull requests welcome.

Page with the same guide: https://accessscan.onetagai.com

MIT License.
