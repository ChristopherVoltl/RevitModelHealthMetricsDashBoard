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
        public string CreateJsonPayload(int totalElements, int warningCount, Dictionary<string, int> familyInstances)
        {
            var json = new
            {
                TotalElements = totalElements,
                Warnings = warningCount,
                Families = familyInstances
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
            string jsonPayload = CreateJsonPayload(totalElements, warningCount, familyInstances);

            // Send the payload to Power BI
            PowerBIPush powerBIPush = new PowerBIPush();
            Task.Run(() => powerBIPush.PushDataToPowerBI(jsonPayload)).Wait();

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
        // PowerBI API URL
        private static readonly string apiUrl = "https://api.powerbi.com/beta/e66e77b4-5724-44d7-8721-06df160450ce/datasets/784e3d0c-eda4-431a-bf73-c49de06cf829/rows?experience=power-bi&key=Sjy3QOn2VB7ZbDXUJYRUEfhEC3vwTnB9VDHMRcn4p8IViv%2B2eE73IBRfzBUDmYvyiLqPNe%2FayWra%2FeYQ1hKHyg%3D%3D";

        public async Task PushDataToPowerBI(string jsonPayload)
        {
            using (HttpClient client = new HttpClient())
            {
                var content = new StringContent(jsonPayload, Encoding.UTF8, "application/json");
                HttpResponseMessage response = await client.PostAsync(apiUrl, content);

                if (!response.IsSuccessStatusCode)
                {
                    // Handle errors here
                    throw new Exception("Failed to post data to PowerBI");
                }
            }
        }
    }
}