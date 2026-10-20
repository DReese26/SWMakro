# SWMakro

## Assembly / Excel synchronization

`SolidWorksExcelSync.bas` is a SolidWorks VBA macro module. It synchronizes the
active, saved assembly with an `.xlsx` workbook:

- Parts in the assembly are listed with their component name, file path,
  matching drawing path (same filename and folder), and configuration.
- A row with a **Part Number** but no **Component** is inserted into the
  assembly. If **Part File** is blank, the macro creates an empty part from the
  configured SolidWorks part template and saves it beside the assembly.
- Existing component rows are refreshed rather than duplicated. The macro does
  not remove components or overwrite existing part files.

### Setup

1. In SolidWorks, create a macro using **Tools > Macro > New**.
2. In the VBA editor, import `SolidWorksExcelSync.bas` using **File > Import
   File**, then save the macro as an `.swp` file.
3. Ensure SolidWorks has a default part template configured. Excel is accessed
   through COM automation, so no Excel reference needs to be added.
4. To run it from a button, use **Tools > Customize > Commands > Macro** and
   assign the saved macro's `main` procedure.

### Use

Run the macro while an assembly is active and saved. Enter the full `.xlsx`
workbook path when prompted; a new workbook is created if the path does not
exist. The first worksheet uses these headers in columns A–E:

| Part Number | Component | Part File | Drawing File | Configuration |
| --- | --- | --- | --- | --- |

Add a part by entering its part number in column A and leaving column B blank.
Leave column C blank to create a new empty part, or provide the full path to an
existing `.SLDPRT` file to insert it. New components are placed at the assembly
origin. Save the workbook after editing and run the macro again to synchronize
the changes.

The macro leaves Excel open. Failed Excel rows are skipped and reported in the
VBA Immediate window; successfully imported parts and assembly changes are
saved.

## SOLIDWORKS PDM drawing publish task

`PdmDrawingPublish.ps1` is a PowerShell worker for a PDM task that runs once for
each selected drawing. It accepts the drawing's local vault-view path, its main
revision, and the vault name:

```powershell
PdmDrawingPublish.ps1 -DrawingPath <local-vault-path-to-SLDDRW> `
  -Revision <PDM-main-revision> -VaultName <vault-name>
```

The PDM task must pass its per-file path and main-revision values as these
arguments. Configure the task to run once for each selected file and supply the
PDM variable that contains the main revision. The worker performs these exports
in separate SOLIDWORKS sessions, closing SOLIDWORKS between exports:

- One all-sheets PDF named `<drawing>-<revision>.pdf` is created in the drawing's
  vault folder and added to the vault or checked in as a new version. A copy is
  also written to `\\srvvcad\AutoCAD\Zeichnungen\`.
- Each drawing sheet is exported as
  `<drawing>-<revision>-<sheet>.dxf` to that network folder.
- The referenced model is exported as `<drawing>-<revision>.step` to that
  network folder. If the drawing references more than one distinct model or
  configuration, the task reports an error rather than choosing one silently.

Run the task on a Windows task host with SOLIDWORKS and the SOLIDWORKS PDM
client/API installed. The task account needs read access to the drawing and
referenced model, permission to add/update the PDF in the vault, and write
access to the network destination. The PDM task configuration must provide
the vault name and the exact main revision for each input drawing. Errors are
reported to the task host's PowerShell error output.
