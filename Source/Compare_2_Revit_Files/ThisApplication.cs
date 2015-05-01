/*
 * Created by SharpDevelop.
 * User: vanyog
 * Date: 4/30/2015
 * Time: 9:11 AM
 * 
 * To change this template use Tools | Options | Coding | Edit Standard Headers.
 */
using System;
using System.Text;
using Autodesk.Revit.UI;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI.Selection;
using System.Collections.Generic;
using System.Linq;

namespace Compare_2_Revit_Files
{
	// Class representing the result of comparison
	public class CpomtareResult
	{
		public double result = 0;
		public double total = 0;
		public string message = "";
		
		public void Compare(string s, Object o1, Object o2){
			if (o1.Equals(o2)) result += 1;
			total += 1;
			message += s + ": " + result.ToString() + "/" + total.ToString() + "\r\n";
		}
	}
	
    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    [Autodesk.Revit.DB.Macros.AddInId("DE3F8111-2108-4DB5-A36C-04AAF7AFE86B")]
	public partial class ThisApplication
	{
		private void Module_Startup(object sender, EventArgs e)
		{

		}

		private void Module_Shutdown(object sender, EventArgs e)
		{

		}

		// Helper function for showing messages
		public void ShowMessage(string m, params Object[] p){
			TaskDialog.Show("Compare 2 projects macro", string.Format(m, p) );
		}
		
		#region Revit Macros generated code
		private void InternalStartup()
		{
			this.Startup += new System.EventHandler(Module_Startup);
			this.Shutdown += new System.EventHandler(Module_Shutdown);
		}
		#endregion
		
		// A string that shows the name and value of a parameter
		public string ParamToString(Parameter p){
			string v = "";
			switch (p.StorageType){
				case StorageType.Double: v = p.AsDouble().ToString(); break;
//				case StorageType.ElementId: v = p.AsElementId().ToString(); break;
				case StorageType.Integer: v = p.AsInteger().ToString(); break;
				case StorageType.String: v = p.AsString(); break;
				default: v = p.StorageType.ToString(); break;
			}
			return p.Definition.Name + "= " + v;
		}
		
		// A string that contains a list of all parameters and their values of an element
		public string ParamsToString(Element e){
			StringBuilder sb = new StringBuilder();
			foreach(Parameter p in e.Parameters){
				sb.AppendLine(ParamToString(p));
			}
			return sb.ToString();
		}
		
		// Dictionary of parameters of an element by names
		public Dictionary<string,Parameter> ParametersOf(Element e){
			Dictionary<string,Parameter> d = new Dictionary<string, Parameter>();
			foreach(Parameter p in e.Parameters){
				d[p.Definition.Name] = p;
			}
			return d;
		}
		
		// List of elements in document d which are of type t
		public List<Element> ElementsOfType(Document d, Type t){
			FilteredElementCollector collector = new FilteredElementCollector(d);
			return collector.OfClass(t).ToList();
		}
		
		// Compares elements of type t from documents d1 and d2 upon a list of parameter names
		public CpomtareResult CompareElementsOtType(Document d1, Document d2, Type t, CpomtareResult rz, params string[] pn){
			List<Element> le1 = ElementsOfType(d1,t);
			Dictionary<string,Parameter> pd1 = ParametersOf(le1[0]);
			rz.message += ParamsToString(le1[0]);
			return rz;
		}
		
		public void Compare2Projects()
		{
			// Checking the count of opened files. Must be 2 to continue.
			int docCount = Application.Documents.Size;
			if (docCount != 2){
				ShowMessage("There must be 2 projects opened to compare them");
				return;
			}
			
			// Full path to the base project we compare
			string baseProjectPathName = @"C:\Users\vanyog\Documents\Project1.rvt";
			
			// Getting the tow files we compare
			Document baseProject = null;
			Document secondProject = null;
			foreach(Document d in Application.Documents){
				if (d.PathName == baseProjectPathName) baseProject = d;
				else secondProject = d;
			}
			
			// Stop if the base project is not opened
			if (baseProject == null){
				ShowMessage("The base project \"{0}\" is not opened.",baseProjectPathName);
				return;
			}
			
			// Compare Project units
			CpomtareResult crz = new CpomtareResult();
			crz.Compare(
				"Units",
				baseProject.GetUnits().GetFormatOptions(UnitType.UT_Length).DisplayUnits,
				secondProject.GetUnits().GetFormatOptions(UnitType.UT_Length).DisplayUnits
			);
			
			CompareElementsOtType(baseProject,secondProject,typeof(Level),crz);
			
			ShowMessage(crz.message);
		}
	}
}