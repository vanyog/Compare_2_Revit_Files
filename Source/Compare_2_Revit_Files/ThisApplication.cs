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
	public class CompareResult
	{
		public double result = 0;
		public double total = 0;
		private StringBuilder sb = new StringBuilder();
		
		public string message{
			get { return sb.ToString(); }
		}
		
		public void AppendLine(string s){
			sb.AppendLine(s);
		}
		
		public void Append(string s){
			sb.Append(s);
		}
		
		public void Compare(string s, Object o1, Object o2){
			if (o1.Equals(o2)) result += 1;
			total += 1;
			sb.AppendLine( s + ": " + result.ToString() + "/" + total.ToString() );
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
		
		// Helper function showing messages limited times
		private static int messCount = 3; // Limit
		public void ShowMessageN(string m, params Object[] p){
			if ( messCount < 1 ) return;
			messCount--;
			ShowMessage(m,p);
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
		
		// List of elements in document d which are of type t
		public List<Element> ElementsOfType(Document d, Type t){
			FilteredElementCollector collector = new FilteredElementCollector(d);
			return collector.OfClass(t).ToList();
		}
		
		// First element in document d which is of type t
		public Element FirstElementOfTyle(Document d, Type t){
			List<Element> le = ElementsOfType(d, t);
			return le[0];
		}
		
		// Dictionary of parameters of an element by names
		public Dictionary<string,Parameter> ParametersOf(Element e){
			Dictionary<string,Parameter> d = new Dictionary<string, Parameter>();
			foreach(Parameter p in e.Parameters){
				d[p.Definition.Name] = p;
			}
			return d;
		}
		
		// Compare 2 parameters
		public Double tolerance = 0.001;
		
		public CompareResult Compare2Params(string n, Parameter p1, Parameter p2, CompareResult rz){
			if (p1.StorageType != p2.StorageType ){
				throw(new Exception("Different parameter storage types can't be compared"));
			}
			Double d = 0;
			string v = "";
			switch (p1.StorageType){
					case StorageType.Double:
					v = p1.AsDouble().ToString() + "-" + p2.AsDouble().ToString();
						if (p1.AsDouble()==0) {
							if (p2.AsDouble()<tolerance) d = 1;
							else d = 0;
						}
						else{
							d = Math.Abs( (p2.AsDouble()-p1.AsDouble()) / p1.AsDouble() );
							if (d<tolerance) d = 1;
							else d = 1 - d;
						}
						break;
					default:
						v = p1.AsString() + "-" + p2.AsString();
						if (p1.AsValueString() == p2.AsValueString()) d = 1;
						else d = 0;
						break;
			}
			rz.result += d;
			rz.total += 1;
			rz.AppendLine( n + ": " + v + "  " + d.ToString() + "/1" );
			return rz;
		}
		// Compare 2 elements upon a list of parameter names
		public CompareResult Compare2Elements(Element e1, Element e2, params string[] pn){
			CompareResult rz = new CompareResult();
			Dictionary<string,Parameter> dp1 = ParametersOf(e1);
			Dictionary<string,Parameter> dp2 = ParametersOf(e2);
			foreach(string n in pn){
				rz = Compare2Params(n,dp1[n],dp2[n],rz);
			}
			return rz;
		}
		
		// Compares elements of type t from documents d1 and d2 upon a list of parameter names
		public CompareResult CompareElementsOtType(Document d1, Document d2, Type t, CompareResult rz, params string[] pn){
			List<Element> le1 = ElementsOfType(d1,t);
			List<Element> le2 = ElementsOfType(d2,t);
			int k = 0; // Number of identical elements found
			for(int i=0; i<le1.Count; i++){
				int m = k;
				CompareResult r0 = Compare2Elements(le1[i],le2[m],pn);
				for(int j=k+1; j<le2.Count; j++){
					CompareResult r1 = Compare2Elements(le1[j],le2[m],pn);
					if (r1.result > r0.result ){
						m = j;
						r0 = r1;
					}
				}
				if (m != k){
					Element e = le2[k];
					le2[k] = le2[m];
					le2[m] = e;
				}
				if (r0.result == r0.total) k++;
				rz.result += r0.result;
				rz.total += r0.total;
				rz.Append(r0.message);
			}
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
			string baseProjectPathName = @"C:\Users\Vanyog\Documents\Project1.rvt";
			
			// Getting the tow files we compare
			Document baseProject = null;
			Document secondProject = null;
			foreach(Document d in Application.Documents){
				if (d.PathName.ToLower() == baseProjectPathName.ToLower()) baseProject = d;
				else secondProject = d;
			}
			
			// Stop if the base project is not opened
			if (baseProject == null){
				ShowMessage("The base project \"{0}\" is not opened.",baseProjectPathName);
				return;
			}
			
			// Compare Project units
			CompareResult crz = new CompareResult();
			crz.Compare(
				"Units",
				baseProject.GetUnits().GetFormatOptions(UnitType.UT_Length).DisplayUnits,
				secondProject.GetUnits().GetFormatOptions(UnitType.UT_Length).DisplayUnits
			);
			
			// Compare tow projects by element types and their main properties
			CompareElementsOtType(baseProject,secondProject,typeof(Level),crz,"Name","Elevation");
			
			// Displaying of the result
			ShowMessage(crz.message);
			
			ShowMessage(ParamsToString(FirstElementOfTyle(baseProject,typeof(Level))));
		}
	}
}