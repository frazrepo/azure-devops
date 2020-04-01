using OfficeOpenXml;
using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
//Epplus library


namespace GetProjectInfos
{
    public static class ExcelSerializer
    {

        public static void Serialize(List<ProjectExtractInfo> projects, List<BuildExtractInfo> builds, List<ReleaseExtractInfo> releases)
        {
            ExcelPackage excel = new ExcelPackage();

            WriteProjectDatas(projects, excel);

            WriteBuildDatas(builds, excel);

            WriteReleaseDatas(releases, excel);

            WriteToFile(excel);

            //Close Excel package 
            excel.Dispose();
        }

        private static void WriteToFile(ExcelPackage excel)
        {
            //Write to File
            string xlPath = "DevOpsProjects.xlsx";

            if (File.Exists(xlPath))
                File.Delete(xlPath);

            // Create excel file on physical disk  
            FileStream objFileStrm = File.Create(xlPath);
            objFileStrm.Close();

            // Write content to excel file  
            File.WriteAllBytes(xlPath, excel.GetAsByteArray());
        }

        private static void WriteProjectDatas(List<ProjectExtractInfo> projects, ExcelPackage excel)
        {
            // Project worksheet
            var workSheet = excel.Workbook.Worksheets.Add("Projects");

            // Header
            workSheet.Row(1).Height = 20;
            workSheet.Row(1).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            workSheet.Row(1).Style.Font.Bold = true;

            // Header of the Excel sheet 
            workSheet.Cells[1, 1].Value = "Name";
            workSheet.Cells[1, 2].Value = "BuildCount";
            workSheet.Cells[1, 3].Value = "ReleaseCount";

            //Projects Content
            int recordIndex = 2;
            for (int i = 0; i < projects.Count; i++)
            {
                workSheet.Cells[recordIndex, 1].Value = projects[i].Name;
                workSheet.Cells[recordIndex, 2].Value = projects[i].BuildCount;
                workSheet.Cells[recordIndex, 3].Value = projects[i].ReleaseCount;

                recordIndex++;
            }

            workSheet.Cells[workSheet.Dimension.Address].AutoFitColumns();

        }

        private static void WriteBuildDatas(List<BuildExtractInfo> builds, ExcelPackage excel)
        {
            // Project worksheet
            var workSheet = excel.Workbook.Worksheets.Add("Builds");

            // Header
            workSheet.Row(1).Height = 20;
            workSheet.Row(1).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            workSheet.Row(1).Style.Font.Bold = true;

            // Header of the Excel sheet 
            workSheet.Cells[1, 1].Value = "Project";
            workSheet.Cells[1, 2].Value = "Build Name";
            workSheet.Cells[1, 3].Value = "RepoType";
            workSheet.Cells[1, 4].Value = "RepoName";
            workSheet.Cells[1, 5].Value = "TriggerType";
            workSheet.Cells[1, 6].Value = "TriggerBranchFilters";
            workSheet.Cells[1, 7].Value = "DefaultAgent";
            workSheet.Cells[1, 8].Value = "Steps";
            workSheet.Cells[1, 9].Value = "Url";

        //Projects Content
        int recordIndex = 2;
            for (int i = 0; i < builds.Count; i++)
            {
                try
                {
                    workSheet.Cells[recordIndex, 1].Value = builds[i].ProjectName;
                    workSheet.Cells[recordIndex, 2].Value = builds[i].Name;
                    workSheet.Cells[recordIndex, 3].Value = builds[i].RepoType;
                    workSheet.Cells[recordIndex, 4].Value = builds[i].RepoName;
                    workSheet.Cells[recordIndex, 5].Value = builds[i].TriggerType;
                    workSheet.Cells[recordIndex, 6].Value = builds[i].TriggerBranchFilters;
                    workSheet.Cells[recordIndex, 7].Value = builds[i].DefaultAgent;
                    workSheet.Cells[recordIndex, 8].Value = builds[i].Steps.Count.ToString();
                    workSheet.Cells[recordIndex, 9].Value = builds[i].Url;
                }
                catch (Exception e)
                {
                 //   Console.WriteLine(e.Message);
                }

                recordIndex++;
            }

            workSheet.Cells[workSheet.Dimension.Address].AutoFitColumns();
        }
        private static void WriteReleaseDatas(List<ReleaseExtractInfo> releases, ExcelPackage excel)
        {
            // Project worksheet
            var workSheet = excel.Workbook.Worksheets.Add("Releases");

            // Header
            workSheet.Row(1).Height = 20;
            workSheet.Row(1).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            workSheet.Row(1).Style.Font.Bold = true;

            // Header of the Excel sheet 
            workSheet.Cells[1, 1].Value = "Project";
            workSheet.Cells[1, 2].Value = "Release Name";

            //Projects Content
            int recordIndex = 2;
            for (int i = 0; i < releases.Count; i++)
            {
                workSheet.Cells[recordIndex, 1].Value = releases[i].ProjectName;
                workSheet.Cells[recordIndex, 2].Value = releases[i].Name;

                recordIndex++;
            }

            workSheet.Cells[workSheet.Dimension.Address].AutoFitColumns();

        }

    }
}
