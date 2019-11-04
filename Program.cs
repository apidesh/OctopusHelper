using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using Octopus.Client;
using Octopus.Client.Model;

namespace OctopusHelper
{
    internal class Program
    {
        private static void Main(string[] args)
        {
            try
            {
                var parametersDictionary = new Dictionary<ParameterEnum, string>();
                if (args.Length == 4)
                {
                    parametersDictionary.Add(ParameterEnum.ServerUrl, args[0].Trim());
                    parametersDictionary.Add(ParameterEnum.ApiKey, args[1].Trim());
                    parametersDictionary.Add(ParameterEnum.ProjectName, args[2].Trim());
                    parametersDictionary.Add(ParameterEnum.OutputPath, args[3].Trim());
                }
                else
                {
                    Console.WriteLine($"Required the parameter of ServerUrl, ApiKey and ProjectName");
                    Environment.Exit(1);
                }

                var endpoint = new OctopusServerEndpoint(parametersDictionary[ParameterEnum.ServerUrl], parametersDictionary[ParameterEnum.ApiKey]);
                var repository = new OctopusRepository(endpoint);

                var project = repository.Projects.FindOne(p => p.Name == parametersDictionary[ParameterEnum.ProjectName]);

                var variablesList = new List<VariableViewModel>();

                //Dictionary to get Names from Ids
                var scopeNames = repository.Environments.FindAll().ToDictionary(x => x.Id, x => x.Name);
                repository.Machines.FindAll().ToList().ForEach(x => scopeNames[x.Id] = x.Name);
                repository.Projects.GetChannels(project).Items.ToList().ForEach(x => scopeNames[x.Id] = x.Name);

                var allEnvironments = repository.Environments.FindAll();

                var allRoles = repository.MachineRoles.GetAllRoleNames();

                var deploymentSteps = repository.DeploymentProcesses.Get(project.DeploymentProcessId).Steps.ToList();

                deploymentSteps.SelectMany(x => x.Actions).ToList().ForEach(x => scopeNames[x.Id] = x.Name);

                //Get All Library Set Variables
                var allVariables = repository.LibraryVariableSets.FindAll();

                ExportVariableSets(parametersDictionary, deploymentSteps, repository, project, variablesList, scopeNames, allVariables, allEnvironments);

                //var input = Console.ReadLine();
                Environment.Exit(0);
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Error {ex.Message}|{ex.StackTrace}");
                Environment.Exit(1);
                throw ex;
            }
        }

        private static void ExportVariableSets(Dictionary<ParameterEnum, string> parametersDictionary, List<DeploymentStepResource> deploymentSteps, OctopusRepository repository, ProjectResource project, List<VariableViewModel> variablesList, Dictionary<string, string> scopeNames, List<LibraryVariableSetResource> allVariables, List<EnvironmentResource> allEnvironments)
        {
            var librarySets = new List<LibraryVariableSetResource>();

            foreach (var item in project.IncludedLibraryVariableSetIds)
            {
                var list = allVariables.FindAll(p => p.Id == item);
                librarySets.AddRange(list);
            }

            foreach (var libraryVariableSetResource in librarySets)
            {
                var variables = repository.VariableSets.Get(libraryVariableSetResource.VariableSetId);
                var variableSetName = libraryVariableSetResource.Name;
                foreach (var variable in variables.Variables)
                {
                    variablesList.Add(new VariableViewModel(variable, variableSetName, scopeNames));
                }
            }

            var dataTableVariable = new DataTable();

            dataTableVariable.Columns.Add("VariableSet", typeof(string));
            dataTableVariable.Columns.Add("Name", typeof(string));
            dataTableVariable.Columns.Add("Value", typeof(string));
            dataTableVariable.Columns.Add("Scope", typeof(string));

            var dataTableDeployStep = new DataTable();

            dataTableDeployStep.Columns.Add("Process", typeof(string));
            dataTableDeployStep.Columns.Add("StepName", typeof(string));
            dataTableDeployStep.Columns.Add("PackageId", typeof(string));
            dataTableDeployStep.Columns.Add("InstalledPath", typeof(string));
            dataTableDeployStep.Columns.Add("TargetRole", typeof(string));
            dataTableDeployStep.Columns.Add("Environment", typeof(string));
            dataTableDeployStep.Columns.Add("ScriptSyntax", typeof(string));
            dataTableDeployStep.Columns.Add("ScriptSource", typeof(string));
            dataTableDeployStep.Columns.Add("ScriptBody", typeof(string));

            //Get All Project Variables for the Project
            var projectSets = repository.VariableSets.Get(project.VariableSetId);

            foreach (var variable in projectSets.Variables)
            {
                variablesList.Add(new VariableViewModel(variable, parametersDictionary[ParameterEnum.ProjectName], scopeNames));
            }

            foreach (var vm in variablesList)
            {
                dataTableVariable.Rows.Add(vm.VariableSetName, vm.Name, vm.Value, vm.Scope);
            }

            foreach (var item in deploymentSteps)
            {
                var process = item.Name;
                var targetRole = item.Properties["Octopus.Action.TargetRoles"].Value;
                foreach (var action in item.Actions)
                {
                    var stepName = action.Name;
                    action.Properties.TryGetValue("Octopus.Action.Package.PackageId", out PropertyValueResource packageId);
                    action.Properties.TryGetValue("Octopus.Action.Package.CustomInstallationDirectory", out PropertyValueResource installedPath);

                    var environment = allEnvironments.FirstOrDefault(p => p.Id == action.Environments.FirstOrDefault())?.Name;
                    action.Properties.TryGetValue("Octopus.Action.Script.Syntax", out PropertyValueResource scriptSyntax);
                    action.Properties.TryGetValue("Octopus.Action.Script.ScriptSource", out PropertyValueResource scriptSource);
                    action.Properties.TryGetValue("Octopus.Action.Script.ScriptBody", out PropertyValueResource scriptBody);

                    dataTableDeployStep.Rows.Add(process, stepName, packageId?.Value, installedPath?.Value, targetRole, environment, scriptSyntax?.Value, scriptSource?.Value, scriptBody?.Value);
                }
            }

            WriteExcelWithNPOI(dataTableVariable, dataTableDeployStep, "xlsx", parametersDictionary[ParameterEnum.ProjectName], parametersDictionary[ParameterEnum.OutputPath]);

            dataTableVariable.Dispose();
            dataTableDeployStep.Dispose();
        }

        public static void WriteExcelWithNPOI(DataTable dtVariables, DataTable dtProcessSteps, string extension, string fileName, string outputPath)
        {
            IWorkbook workbook;

            Console.WriteLine($"Export data with {extension} file extension");

            if (extension == "xlsx")
            {
                workbook = new XSSFWorkbook();
            }
            else if (extension == "xls")
            {
                workbook = new HSSFWorkbook();
            }
            else
            {
                throw new Exception("This format is not supported");
            }

            ISheet librarySetSheet = workbook.CreateSheet("Library Sets");
            ISheet deploymentProcessSheet = workbook.CreateSheet("Deployment Process");

            CreateDataSheet(dtVariables, librarySetSheet);
            CreateDataSheet(dtProcessSteps, deploymentProcessSheet);

            if (!Directory.Exists(outputPath))
            {
                Directory.CreateDirectory(outputPath);
            }

            var outputFolder = DateTime.Now.ToString("yyyyMMdd");
            var targetPath = Path.Combine(outputPath, outputFolder);
            if (!Directory.Exists(targetPath))
            {
                Directory.CreateDirectory(targetPath);
            }

            var outputFileName = $"{fileName.Replace(" ", "_")}_{ DateTime.Now.ToString("yyyyMMdd_HHmmss")}.{ extension}";
            var outputFileInfo = Path.Combine(targetPath, outputFileName);
            Console.WriteLine($"Success to export data {outputFileInfo}");
            using (var fs = File.Create(outputFileInfo))
            {
                workbook.Write(fs);
            }
        }

        private static void CreateDataSheet(DataTable dt, ISheet sheet)
        {
            //make a header row
            IRow rowHeader = sheet.CreateRow(0);

            for (int j = 0; j < dt.Columns.Count; j++)
            {
                ICell cell = rowHeader.CreateCell(j);
                string columnName = dt.Columns[j].ToString();
                cell.SetCellValue(columnName);
            }

            //loops through data
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                IRow rowData = sheet.CreateRow(i + 1);
                for (int j = 0; j < dt.Columns.Count; j++)
                {
                    ICell cell = rowData.CreateCell(j);
                    string columnName = dt.Columns[j].ToString();
                    cell.SetCellValue(dt.Rows[i][columnName].ToString());
                }
            }
        }
    }
}