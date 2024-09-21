using System;
using System.Collections.Generic;
using System.IO;
using System.Text.Json;
using INFITF;
using MECMOD;
using ProductStructureTypeLib;
using SPATypeLib; // For geometric properties like material

namespace SimpleCatiaExtractor
{
    public class CATIAProductExtractor
    {
        public void ExtractProductProperties(string productFilePath, string outputJsonFilePath)
        {
            try
            {
                // Start CATIA Application
                Application catiaApp = (Application)Activator.CreateInstance(Type.GetTypeFromProgID("CATIA.Application"));
                ProductDocument productDocument = (ProductDocument)catiaApp.Documents.Open(productFilePath);
                Products products = productDocument.Product.Products;

                // Store extracted part data
                List<object> partDataList = new List<object>();

                // Loop through each part in the product
                for (int i = 1; i <= products.Count; i++)
                {
                    Product part = products.Item(i);
                    Console.WriteLine($"Extracting properties for Part: {part.get_Name()}");

                    // Extract part properties
                    var partData = new
                    {
                        Name = part.get_Name(),
                        Features = GetFeatures(part),
                        Dimensions = GetBoundingBox(part),
                        Material = GetMaterial(part),
                        Color = GetColor(part)
                    };

                    partDataList.Add(partData);
                }

                // Save extracted properties to JSON file
                string jsonString = JsonSerializer.Serialize(partDataList, new JsonSerializerOptions { WriteIndented = true });
                System.IO.File.WriteAllText(outputJsonFilePath, jsonString);

                // Close the product document after extraction
                productDocument.Close();
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error occurred: " + ex.Message);
            }
        }

        // Helper method to get features of the part
        public List<string> GetFeatures(Product part)
        {
            List<string> features = new List<string>();

            try
            {
                // Access the part's bodies (simple example)
                Bodies bodies = part.ReferenceProduct.Parent as Bodies;
                if (bodies != null)
                {
                    for (int i = 1; i <= bodies.Count; i++)
                    {
                        Body body = bodies.Item(i);
                        features.Add($"Body: {body.get_Name()}");

                        // Example of adding shapes from the body (simplified)
                        foreach (Shape shape in body.Shapes)
                        {
                            features.Add($"Shape: {shape.get_Name()}");
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error extracting features: " + ex.Message);
            }

            return features;
        }

        // Helper method to get bounding box (min/max points) for the part
        public object GetBoundingBox(Product part)
        {
            try
            {
                // Placeholder for min/max points (bounding box)
                double[] minPoint = { double.MaxValue, double.MaxValue, double.MaxValue };
                double[] maxPoint = { double.MinValue, double.MinValue, double.MinValue };

                // Example: Iterate through part geometry (this is simplified)
                // Here, you would access the part's bodies, shapes, or geometry
                Bodies bodies = part.ReferenceProduct.Parent as Bodies;
                if (bodies != null)
                {
                    for (int i = 1; i <= bodies.Count; i++)
                    {
                        Body body = bodies.Item(i);

                        // Process each shape to find min/max coordinates
                        foreach (Shape shape in body.Shapes)
                        {
                            // Access shape's geometry to compute bounding box
                            // (You would need to calculate real geometry here)
                            // Placeholder logic:
                            UpdateBoundingBox(minPoint, maxPoint, shape);
                        }
                    }
                }

                // Return the bounding box data
                return new
                {
                    MinX = minPoint[0],
                    MinY = minPoint[1],
                    MinZ = minPoint[2],
                    MaxX = maxPoint[0],
                    MaxY = maxPoint[1],
                    MaxZ = maxPoint[2]
                };
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error calculating bounding box: " + ex.Message);
                return new { MinX = 0.0, MinY = 0.0, MinZ = 0.0, MaxX = 0.0, MaxY = 0.0, MaxZ = 0.0 };
            }
        }

        // Helper method to update the bounding box coordinates
        public void UpdateBoundingBox(double[] minPoint, double[] maxPoint, Shape shape)
        {
            // Placeholder logic to update min/max points based on the shape's geometry
            // (Replace with real geometric calculations)
            minPoint[0] = Math.Min(minPoint[0], 0);  // Update based on actual coordinates
            maxPoint[0] = Math.Max(maxPoint[0], 100); // Placeholder max
            // Repeat for Y and Z
        }

        // Method to get material of the part
        public string GetMaterial(Product part)
        {
            try
            {
                MaterialManager materialManager = part.MaterialManager as MaterialManager;
                if (materialManager != null && materialManager.Materials.Count > 0)
                {
                    Material material = materialManager.Materials.Item(1);  // Getting the first material
                    return material.Name;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error retrieving material: " + ex.Message);
            }
            return "No Material";
        }

        // Method to get color of the part
        public string GetColor(Product part)
        {
            try
            {
                var visProperties = (SPAWorkbench)part.ReferenceProduct.Parent;
                if (visProperties != null)
                {
                    // Simplified: Extract color-related properties (e.g., RGB values)
                    // Placeholder for getting the actual color
                    return "Red";  // Replace with actual color extraction logic
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error retrieving color: " + ex.Message);
            }
            return "No Color";
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            // Input and output paths
            string productFilePath = @"C:\path_to_catia_product_file.CATProduct";
            string outputJsonFilePath = @"product_properties_test.json";

            // Create an instance of the extractor and run the extraction
            CATIAProductExtractor extractor = new CATIAProductExtractor();
            extractor.ExtractProductProperties(productFilePath, outputJsonFilePath);

            Console.WriteLine("Extraction completed. Properties saved to JSON.");
            Console.ReadKey();
        }
    }
}
