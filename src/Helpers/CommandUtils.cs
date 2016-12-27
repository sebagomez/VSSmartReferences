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
		public static IEnumerable<object> GetSelectedReferences(DTE2 dte)
		{
			var selectedItems = (Array)dte.ToolWindows.SolutionExplorer.SelectedItems;
			return selectedItems.Cast<UIHierarchyItem>().Select(x => x.Object);
		}
	}
}
