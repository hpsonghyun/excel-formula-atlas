# Security and workbook privacy

The analyzer reads the selected workbook and writes a self-contained HTML snapshot. It does not upload the workbook to a remote service. The HTML includes collected cell values, formulas, nearby labels and worksheet names. Treat a generated map as workbook data when sharing it.

Run VBA only in a workbook and environment you trust. The optional connection to desktop Excel uses an embedded PowerShell helper on `127.0.0.1` and temporary local session files. It is limited to selecting captured workbook locations. Follow your organization's macro and PowerShell policies. The analyzer does not change those policies.

A newly selected workbook is opened read-only with link updates disabled and VBA automation macros disabled. Excel may still apply its own security prompts for other legacy macro types. An already open workbook stays under its existing Excel session settings.

The analyzer does not write source cell values or save the source workbook. Navigation can activate the source workbook and select a cell. Use `ExcelMap_Stop` before closing the macro host or saving it under another name. Rebuild the map after changing source values, formulas or paths.

For a bug report, use the synthetic example or a minimal workbook with made-up data. Do not attach confidential workbooks, generated maps, screenshots, temporary connection files or authentication material to a public issue. A private reporting contact will be configured when the repository is published.
