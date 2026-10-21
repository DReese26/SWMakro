# SWMakro

## SOLIDWORKS PDM drawing export task

`SWMakro.PdmTask` contains a C# task add-in for SOLIDWORKS PDM Professional 2022. It processes every selected SLDDRW sequentially and creates:

- One PDF containing all drawing sheets.
- One DXF per sheet, named `<drawing>_<sheet>.dxf`.
- One STEP file named `<drawing>.step`, exported from the single referenced part or assembly and the configuration used by its drawing views.

Outputs are placed beside the source drawing in the vault. New outputs are added to PDM; existing outputs are checked out, replaced, and checked in. A source drawing checked out by a user, an output checked out by another user, unavailable references, missing configurations, and drawings referencing multiple distinct model/configuration pairs are reported as task failures rather than silently exporting an arbitrary model.

### Build and install

Build `SWMakro.PdmTask/SWMakro.PdmTask.csproj` on a Windows machine with .NET Framework 4.8, SOLIDWORKS 2022, and the PDM Professional 2022 client installed. The project references the interop assemblies in their default installation locations; adjust the hint paths if your installation uses different folders. Register the compiled assembly as a PDM add-in using the PDM Administration tool, then create a task using **SWMakro PDF-DXF-STEP** and enable its selection menu command. Configure a task host with SOLIDWORKS 2022 installed, an available SOLIDWORKS license, and access to the vault view. Run the task on the task host interactively during initial validation.

### Validation scenarios

Test a drawing with multiple sheets, a drawing whose views all use the same model/configuration, both part and assembly references, and multiple selected drawings. Also verify handling of checked-out source/output files, missing or unavailable model files, missing configurations, and drawings with multiple distinct referenced model/configuration pairs. Confirm that all generated files appear in the vault and that the PDF has every sheet while each DXF has only its corresponding sheet.