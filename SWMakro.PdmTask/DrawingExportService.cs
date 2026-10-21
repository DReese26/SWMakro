using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using EPDM.Interop.epdm;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;

namespace SWMakro.PdmTask
{
    internal sealed class DrawingExportService
    {
        private readonly ISldWorks _solidWorks;
        private readonly IEdmVault5 _vault;

        public DrawingExportService(ISldWorks solidWorks, IEdmVault5 vault)
        {
            _solidWorks = solidWorks;
            _vault = vault;
        }

        public void ProcessDrawing(string drawingPath)
        {
            if (string.IsNullOrWhiteSpace(drawingPath) || !File.Exists(drawingPath))
            {
                throw new FileNotFoundException("The selected drawing is not available in the task-host vault view.", drawingPath);
            }

            IEdmFolder5 sourceFolder;
            var sourceFile = _vault.GetFileFromPath(drawingPath, out sourceFolder);
            if (sourceFile == null || sourceFolder == null)
            {
                throw new InvalidOperationException("The selected file is not available in the PDM vault.");
            }

            if (sourceFile.IsLocked)
            {
                throw new InvalidOperationException("The drawing is checked out. Check it in before running the export task.");
            }

            var stagingDirectory = Path.Combine(Path.GetTempPath(), "SWMakroPdmTask", Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(stagingDirectory);

            try
            {
                var drawing = OpenDocument(drawingPath, swDocumentTypes_e.swDocDRAWING);
                try
                {
                    var modelReference = FindUniqueModelReference(drawing);
                    var modelPath = modelReference.Path;
                    if (!File.Exists(modelPath))
                    {
                        throw new FileNotFoundException("The model referenced by the drawing is not available on the task host.", modelPath);
                    }

                    var modelType = Path.GetExtension(modelPath).Equals(".sldasm", StringComparison.OrdinalIgnoreCase)
                        ? swDocumentTypes_e.swDocASSEMBLY
                        : swDocumentTypes_e.swDocPART;
                    var model = OpenDocument(modelPath, modelType, modelReference.Configuration);
                    try
                    {
                        ExportPdf(drawing, Path.Combine(stagingDirectory, Path.GetFileNameWithoutExtension(drawingPath) + ".pdf"));
                        ExportSheetDxfs(drawing, drawingPath, stagingDirectory);
                        ExportStep(model, modelReference.Configuration, Path.Combine(stagingDirectory, Path.GetFileNameWithoutExtension(drawingPath) + ".step"));
                    }
                    finally
                    {
                        CloseDocument(model);
                    }
                }
                finally
                {
                    CloseDocument(drawing);
                }

                foreach (var stagedPath in Directory.GetFiles(stagingDirectory))
                {
                    PublishFile(stagedPath, Path.Combine(Path.GetDirectoryName(drawingPath), Path.GetFileName(stagedPath)));
                }
            }
            finally
            {
                if (Directory.Exists(stagingDirectory))
                {
                    Directory.Delete(stagingDirectory, true);
                }
            }
        }

        private ModelReference FindUniqueModelReference(IModelDoc2 drawing)
        {
            var references = new Dictionary<string, ModelReference>(StringComparer.OrdinalIgnoreCase);
            var drawingDocument = (IDrawingDoc)drawing;
            var view = drawingDocument.GetFirstView() as IView;
            if (view != null)
            {
                view = view.GetNextView() as IView;
            }

            while (view != null)
            {
                var path = view.GetReferencedModelName();
                if (!string.IsNullOrWhiteSpace(path))
                {
                    var configuration = view.ReferencedConfiguration;
                    if (string.IsNullOrWhiteSpace(configuration))
                    {
                        throw new InvalidOperationException("A drawing view has no referenced model configuration: " + path);
                    }

                    var fullPath = Path.GetFullPath(path);
                    var key = fullPath + "|" + configuration;
                    references[key] = new ModelReference(fullPath, configuration);
                }

                view = view.GetNextView() as IView;
            }

            if (references.Count == 0)
            {
                throw new InvalidOperationException("No referenced part or assembly was found in the drawing views.");
            }

            if (references.Count > 1)
            {
                throw new InvalidOperationException(
                    "The drawing references more than one model/configuration. A single STEP output would be ambiguous: " +
                    string.Join("; ", references.Values.Select(reference => reference.Path + " [" + reference.Configuration + "]")));
            }

            return references.Values.Single();
        }

        private IModelDoc2 OpenDocument(string path, swDocumentTypes_e documentType, string configuration = "")
        {
            var errors = 0;
            var warnings = 0;
            var document = _solidWorks.OpenDoc6(
                path,
                (int)documentType,
                (int)(swOpenDocOptions_e.swOpenDocOptions_Silent | swOpenDocOptions_e.swOpenDocOptions_ReadOnly),
                configuration ?? string.Empty,
                ref errors,
                ref warnings);

            if (document == null || errors != 0)
            {
                throw new InvalidOperationException("SOLIDWORKS could not open '" + path + "' (error " + errors + ").");
            }

            if (!string.IsNullOrWhiteSpace(configuration) && !document.ShowConfiguration2(configuration))
            {
                throw new InvalidOperationException("The referenced configuration '" + configuration + "' is unavailable in " + path + ".");
            }

            return document;
        }

        private static void ExportPdf(IModelDoc2 drawing, string outputPath)
        {
            var exportData = (IExportPdfData)drawing.GetExportFileData((int)swExportDataFileType_e.swExportPdfData);
            if (exportData == null || !exportData.SetSheets((int)swExportDataSheetsToExport_e.swExportData_ExportAllSheets, null))
            {
                throw new InvalidOperationException("SOLIDWORKS could not configure the all-sheet PDF export.");
            }

            SaveAs(drawing, outputPath, exportData);
        }

        private static void ExportSheetDxfs(IModelDoc2 drawing, string drawingPath, string stagingDirectory)
        {
            var drawingDocument = (IDrawingDoc)drawing;
            var sheetNames = drawingDocument.GetSheetNames() as string[];
            if (sheetNames == null || sheetNames.Length == 0)
            {
                throw new InvalidOperationException("The drawing contains no sheets.");
            }

            var usedNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (var sheetName in sheetNames)
            {
                var safeSheetName = MakeSafeFileName(sheetName);
                if (!usedNames.Add(safeSheetName))
                {
                    throw new InvalidOperationException("Sheet names become duplicate DXF filenames after sanitizing: " + sheetName);
                }

                var exportData = (IExportDxfData)drawing.GetExportFileData((int)swExportDataFileType_e.swExportDxfDwgData);
                if (exportData == null ||
                    !exportData.SetSheets(
                        (int)swExportDataSheetsToExport_e.swExportData_ExportSpecifiedSheets,
                        new[] { sheetName }))
                {
                    throw new InvalidOperationException("SOLIDWORKS could not configure DXF export for sheet '" + sheetName + "'.");
                }

                var outputPath = Path.Combine(
                    stagingDirectory,
                    Path.GetFileNameWithoutExtension(drawingPath) + "_" + safeSheetName + ".dxf");
                SaveAs(drawing, outputPath, exportData);
            }
        }

        private static void ExportStep(IModelDoc2 model, string configuration, string outputPath)
        {
            if (!model.ShowConfiguration2(configuration))
            {
                throw new InvalidOperationException("Could not activate STEP configuration '" + configuration + "'.");
            }

            SaveAs(model, outputPath, null);
        }

        private static void SaveAs(IModelDoc2 document, string outputPath, object exportData)
        {
            var errors = 0;
            var warnings = 0;
            var saved = document.Extension.SaveAs(
                outputPath,
                (int)swSaveAsVersion_e.swSaveAsCurrentVersion,
                (int)swSaveAsOptions_e.swSaveAsOptions_Silent,
                exportData,
                ref errors,
                ref warnings);

            if (!saved || errors != 0 || !File.Exists(outputPath))
            {
                throw new InvalidOperationException("SOLIDWORKS failed to save '" + outputPath + "' (error " + errors + ").");
            }
        }

        private void PublishFile(string stagedPath, string vaultPath)
        {
            IEdmFolder5 folder;
            var existingFile = _vault.GetFileFromPath(vaultPath, out folder);
            if (folder == null)
            {
                folder = _vault.GetFolderFromPath(Path.GetDirectoryName(vaultPath));
            }

            if (folder == null)
            {
                throw new InvalidOperationException("The output folder is not present in the vault: " + Path.GetDirectoryName(vaultPath));
            }

            if (existingFile == null)
            {
                var fileId = folder.AddFile2(
                    0,
                    stagedPath,
                    out var error,
                    "Generated by SWMakro PDF-DXF-STEP task",
                    (int)EdmAddFlag.EdmAdd_UniqueVarDelayCheck);
                if (fileId <= 0 || error != 0)
                {
                    throw new InvalidOperationException("PDM could not add output '" + vaultPath + "' (error " + error + ").");
                }

                return;
            }

            if (existingFile.IsLocked)
            {
                throw new InvalidOperationException("Output file is checked out and cannot be replaced: " + vaultPath);
            }

            existingFile.LockFile(folder.ID, 0, 0);
            try
            {
                File.Copy(stagedPath, vaultPath, true);
                existingFile.UnlockFile(0, "Regenerated by SWMakro PDF-DXF-STEP task", (int)EdmUnlockFlag.EdmUnlock_IgnoreReferences);
            }
            catch
            {
                throw;
            }
        }

        private void CloseDocument(IModelDoc2 document)
        {
            if (document == null)
            {
                return;
            }

            try
            {
                _solidWorks.CloseDoc(document.GetTitle());
            }
            finally
            {
                Marshal.FinalReleaseComObject(document);
            }
        }

        private static string MakeSafeFileName(string name)
        {
            var invalidCharacters = Path.GetInvalidFileNameChars();
            var safeName = new string(name.Select(character =>
                invalidCharacters.Contains(character) || character == '/' || character == '\\' ? '_' : character).ToArray()).Trim().TrimEnd('.');
            return string.IsNullOrWhiteSpace(safeName) ? "Sheet" : safeName;
        }

        private sealed class ModelReference
        {
            public ModelReference(string path, string configuration)
            {
                Path = path;
                Configuration = configuration;
            }

            public string Path { get; }
            public string Configuration { get; }
        }
    }
}
