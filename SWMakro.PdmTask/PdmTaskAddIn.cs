using System;
using System.Runtime.InteropServices;
using EPDM.Interop.epdm;

namespace SWMakro.PdmTask
{
    [ComVisible(true)]
    [Guid("FD6F4A49-2EB0-4B7E-A97D-B66D2DB0B01F")]
    [ClassInterface(ClassInterfaceType.None)]
    public sealed class PdmTaskAddIn : IEdmAddIn5
    {
        private const string TaskName = "SWMakro PDF-DXF-STEP";

        public void GetAddInInfo(ref EdmAddInInfo info, IEdmVault5 vault, IEdmCmdMgr5 commandManager)
        {
            info.mbsAddInName = TaskName;
            info.mbsCompany = "SWMakro";
            info.mbsDescription = "Exports selected SolidWorks drawings as PDF, per-sheet DXF, and referenced-model STEP.";
            info.mlAddInVersion = 1;
            info.mlRequiredVersionMajor = 30;
            info.mlRequiredVersionMinor = 0;

            commandManager.AddHook(EdmCmdType.EdmCmd_TaskSetup);
            commandManager.AddHook(EdmCmdType.EdmCmd_TaskSetupButton);
            commandManager.AddHook(EdmCmdType.EdmCmd_TaskLaunch);
            commandManager.AddHook(EdmCmdType.EdmCmd_TaskRun);
        }

        public void OnCmd(ref EdmCmd command, ref EdmCmdData[] data)
        {
            switch (command.meCmdType)
            {
                case EdmCmdType.EdmCmd_TaskSetup:
                    ConfigureTask(command);
                    break;
                case EdmCmdType.EdmCmd_TaskSetupButton:
                    SaveTaskSetup(command);
                    break;
                case EdmCmdType.EdmCmd_TaskLaunch:
                    break;
                case EdmCmdType.EdmCmd_TaskRun:
                    RunTask(command, data);
                    break;
            }
        }

        private static void ConfigureTask(EdmCmd command)
        {
            var properties = command.mpoExtra as IEdmTaskProperties;
            if (properties != null)
            {
                properties.TaskFlags = (int)EdmTaskFlag.EdmTask_SupportsInitExec;
            }
        }

        private static void SaveTaskSetup(EdmCmd command)
        {
            if (!string.Equals(command.mbsComment, "OK", StringComparison.OrdinalIgnoreCase))
            {
                return;
            }

            var properties = command.mpoExtra as IEdmTaskProperties;
            if (properties == null)
            {
                return;
            }

            var menuCommands = new EdmTaskMenuCmd[1];
            menuCommands[0].mbsMenuString = TaskName;
            menuCommands[0].mbsStatusBarHelp = "Create PDF, per-sheet DXFs, and STEP from the selected drawing(s).";
            menuCommands[0].mlCmdID = 1;
            menuCommands[0].mlEdmMenuFlags = (int)EdmMenuFlags.EdmMenu_MustHaveSelection;
            properties.SetMenuCmds(menuCommands);
        }

        private static void RunTask(EdmCmd command, EdmCmdData[] data)
        {
            var task = command.mpoExtra as IEdmTaskInstance;
            if (task == null)
            {
                return;
            }

            task.SetStatus(EdmTaskStatus.EdmTaskStat_Running);
            var vault = command.mpoVault as IEdmVault5;
            var solidWorks = default(SolidWorks.Interop.sldworks.ISldWorks);
            var succeeded = 0;
            var failed = 0;
            var files = new System.Collections.Generic.List<EdmCmdData>();

            if (data != null)
            {
                foreach (EdmCmdData item in data)
                {
                    if (item.mlLongData1 == (long)EdmObjectType.EdmObject_File)
                    {
                        files.Add(item);
                    }
                }
            }

            task.SetProgressRange(files.Count, 0, "Preparing drawing exports");
            try
            {
                if (vault == null)
                {
                    throw new InvalidOperationException("The PDM vault interface is unavailable.");
                }

                solidWorks = (SolidWorks.Interop.sldworks.ISldWorks)Activator.CreateInstance(
                    Type.GetTypeFromProgID("SldWorks.Application", true));
                solidWorks.Visible = false;
                solidWorks.UserControl = false;

                var exporter = new DrawingExportService(solidWorks, vault);
                for (var index = 0; index < files.Count; index++)
                {
                    var item = files[index];
                    var sourcePath = item.mbsStrData1;
                    try
                    {
                        exporter.ProcessDrawing(sourcePath);
                        succeeded++;
                        task.SetStatusMessage(item.mlObjectID1, "Created PDF, sheet DXFs, and STEP for " + sourcePath);
                    }
                    catch (Exception exception)
                    {
                        failed++;
                        task.SetStatusMessage(item.mlObjectID1, "Export failed for " + sourcePath + ": " + exception.Message);
                    }

                    task.SetProgressPos(index + 1, "Processed " + (index + 1) + " of " + files.Count + " drawings");
                }

                if (files.Count == 0)
                {
                    failed++;
                    task.SetStatusMessage(0, "No drawing files were supplied to the task.");
                }

                task.SetStatus(
                    failed == 0 ? EdmTaskStatus.EdmTaskStat_DoneOK : EdmTaskStatus.EdmTaskStat_DoneFailed,
                    0,
                    succeeded + " drawing(s) exported; " + failed + " failed.");
            }
            catch (Exception exception)
            {
                task.SetStatus(EdmTaskStatus.EdmTaskStat_DoneFailed, 0, exception.Message);
            }
            finally
            {
                if (solidWorks != null)
                {
                    try
                    {
                        solidWorks.ExitApp();
                    }
                    catch
                    {
                    }

                    Marshal.FinalReleaseComObject(solidWorks);
                }
            }
        }
    }
}
