using EnvDTE;
using EnvDTE80;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace VSSmartReferences.Helpers
{
	internal class CommandUtils
	{
		public const string ITEM_GORUP = "ItemGroup";
		public const string CONDITION = "Condition";
		public const string REFERENCE = "Reference";
		public const string INCLUDE = "Include";
		public const string HINTPATH = "HintPath";
		public const string PRIVATE = "Private";


		public const string BUILD_INSIDE_VS = "'$(BuildingInsideVisualStudio)' == 'true'";
		public const string BUILD_OUTSIDE_VS = "'$(BuildingInsideVisualStudio)' != 'true'";

		public static IEnumerable<object> GetSelectedReferences(DTE2 dte)
		{
			var selectedItems = (Array)dte.ToolWindows.SolutionExplorer.SelectedItems;
			return selectedItems.Cast<UIHierarchyItem>().Select(x => x.Object);
		}
	}
}
