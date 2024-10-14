using Autodesk.Revit.UI;
using Autodesk.Revit.DB;
using System.Linq;
using System.Reflection.Metadata;

using System.Net.Http;
using System.Text;
using System.Threading.Tasks;
using Autodesk.Revit.Attributes;
using Newtonsoft.Json;


namespace Modelhealth
{
    [Transaction(TransactionMode.Manual)]

    public class ModelHealthMetrics : IExternalCommand
    {
        public string CreateJsonPayload(int totalElements, int warningCount)
        {
            var json = new
            {     

                TotalElements = totalElements,
                Warnings = warningCount,
                //Families = familyInstances
            };

            return JsonConvert.SerializeObject(json);
        }
        // Implement the Execute method of the IExternalCommand interface
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            // Access the Revit document
            Autodesk.Revit.DB.Document doc = commandData.Application.ActiveUIDocument.Document;

            // Call your health check methods
            int totalElements = GetTotalElements(doc);
            int warningCount = GetWarningCount(doc);
            Dictionary<string, int> familyInstances = GetFamilyInstanceCounts(doc);

            //Show data in a TaskDialog or do additional processing
            TaskDialog.Show("Revit Model Health",
                $"Total Elements: {totalElements}\nWarnings: {warningCount}");

            // Prepare the JSON payload
            string jsonPayload = CreateJsonPayload(totalElements, warningCount);

            // Create and run a task to send the payload to Power BI asynchronously
            Task powerBIPushTask = Task.Run(async () =>
            {
                PowerBIPush powerBIPush = new PowerBIPush();
                await powerBIPush.PushDataToPowerBI(jsonPayload);
            });

            // Wait for the task to complete but avoid blocking the UI thread
            powerBIPushTask.GetAwaiter().GetResult();


            TaskDialog.Show("Revit Model Health", "Data has been successfully sent to Power BI.");

            return Result.Succeeded;
        }

        // Method to get total elements
        public int GetTotalElements(Autodesk.Revit.DB.Document doc)
        {
            FilteredElementCollector collector = new FilteredElementCollector(doc);
            return collector.WhereElementIsNotElementType().Count();
        }

        // Method to get warning count
        public int GetWarningCount(Autodesk.Revit.DB.Document doc)
        {
            IList<FailureMessage> warnings = doc.GetWarnings();
            return warnings.Count;
        }

        // Method to get family instance counts
        public Dictionary<string, int> GetFamilyInstanceCounts(Autodesk.Revit.DB.Document doc)
        {
            FilteredElementCollector collector = new FilteredElementCollector(doc)
                .OfClass(typeof(FamilyInstance));

            var familyCounts = collector
                .GroupBy(f => f.Name)
                .ToDictionary(group => group.Key, group => group.Count());

            return familyCounts;
        }
    }

    
    public class PowerBIPush
    {
        public async Task PushDataToPowerBI(string jsonPayload)
        {
            using (HttpClient client = new HttpClient())
            {
                // PowerBI API URL
                string powerBiApiUrl = "https://api.powerbi.com/beta/e66e77b4-5724-44d7-8721-06df160450ce/datasets/07efce07-a7ee-4998-88a6-3adcb6feb25e/rows?experience=power-bi&key=RSQali93QiMzk53VCFG1OmXKXXvkpDujTYpkM8DfNjxB4o0CZ16uIiFciXYFi9JjYtQ2H0mjz3ztwHksnDQkfQ%3D%3D";

                var content = new StringContent(jsonPayload, Encoding.UTF8, "application/json");

                try
                {
                    HttpResponseMessage response = await client.PostAsync(powerBiApiUrl, content);

                    if (!response.IsSuccessStatusCode)
                    {
                        string responseBody = await response.Content.ReadAsStringAsync();
                        throw new Exception($"Failed to post data to Power BI: {response.StatusCode} {response.ReasonPhrase}. Response body: {responseBody}");
                    }
                }
                catch (HttpRequestException ex)
                {
                    throw new Exception($"Error connecting to Power BI: {ex.Message}");
                }
            }
        }
    }
}