//------------------------------------------------------------------------------
// <copyright file="FixReferenceCommand.cs" company="Company">
//     Copyright (c) Company.  All rights reserved.
// </copyright>
//------------------------------------------------------------------------------

using System;
using System.ComponentModel.Design;
using System.Globalization;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using System.Diagnostics;
using EnvDTE;
using EnvDTE80;
using VSSmartReferences.Helpers;
using System.Xml;

namespace VSSmartReferences
{
	/// <summary>
	/// Command handler
	/// </summary>
	internal sealed class FixReferenceCommand
	{
		/// <summary>
		/// Command ID.
		/// </summary>
		public const int CommandId = 0x0100;

		/// <summary>
		/// Command menu group (command set GUID).
		/// </summary>
		public static readonly Guid CommandSet = new Guid("0387718b-56c5-47c6-8c88-53b48714ca34");

		/// <summary>
		/// VS Package that provides this command, not null.
		/// </summary>
		private readonly Package package;

		/// <summary>
		/// Initializes a new instance of the <see cref="FixReferenceCommand"/> class.
		/// Adds our command handlers for menu (commands must exist in the command table file)
		/// </summary>
		/// <param name="package">Owner package, not null.</param>
		private FixReferenceCommand(Package package)
		{
			if (package == null)
			{
				throw new ArgumentNullException("package");
			}

			this.package = package;

			OleMenuCommandService commandService = this.ServiceProvider.GetService(typeof(IMenuCommandService)) as OleMenuCommandService;
			if (commandService != null)
			{
				var menuCommandID = new CommandID(CommandSet, CommandId);
				var menuItem = new MenuCommand(this.FixReferenceCallback, menuCommandID);
				commandService.AddCommand(menuItem);
			}
		}

		/// <summary>
		/// Gets the instance of the command.
		/// </summary>
		public static FixReferenceCommand Instance
		{
			get;
			private set;
		}

		/// <summary>
		/// Gets the service provider from the owner package.
		/// </summary>
		private IServiceProvider ServiceProvider
		{
			get
			{
				return this.package;
			}
		}

		/// <summary>
		/// Initializes the singleton instance of the command.
		/// </summary>
		/// <param name="package">Owner package, not null.</param>
		public static void Initialize(Package package)
		{
			Instance = new FixReferenceCommand(package);
		}

		/// <summary>
		/// This function is the callback used to execute the command when the menu item is clicked.
		/// See the constructor to see how the menu item is associated with this function using
		/// OleMenuCommandService service and MenuCommand class.
		/// </summary>
		/// <param name="sender">Event sender.</param>
		/// <param name="e">Event args.</param>
		private void MenuItemCallback(object sender, EventArgs e)
		{
			string message = string.Format(CultureInfo.CurrentCulture, "Inside {0}.MenuItemCallback()", this.GetType().FullName);
			string title = "FixReferenceCommand";

			// Show a message box to prove we were here
			VsShellUtilities.ShowMessageBox(
				this.ServiceProvider,
				message,
				title,
				OLEMSGICON.OLEMSGICON_INFO,
				OLEMSGBUTTON.OLEMSGBUTTON_OK,
				OLEMSGDEFBUTTON.OLEMSGDEFBUTTON_FIRST);
		}

		private void FixReferenceCallback(object sender, EventArgs e)
		{
			var dte = (DTE2)ServiceProvider.GetService(typeof(DTE));
			foreach (object item in CommandUtils.GetSelectedReferences(dte))
			{
				var d = (dynamic)item;
				if (d.SourceProject == null)
					continue;

				string libPath = d.Path;
				string sourceProject = d.SourceProject.FullName;
				string currentProject = d.ContainingProject.FullName;


				FixReference(currentProject, libPath, sourceProject);

				//System.Diagnostics.Debug.WriteLine($"CurrentProject: {currentProject}, Path:{libPath}, Project:{sourceProject}");

			}

		}

		void FixReference(string projectFile, string libRef, string projRef)
		{
			XmlDocument xml = new XmlDocument();
			xml.Load(projectFile);
			XmlElement root = xml.DocumentElement;
			XmlNamespaceManager nsMgr = new XmlNamespaceManager(xml.NameTable);
			nsMgr.AddNamespace("ns", "http://schemas.microsoft.com/developer/msbuild/2003");

			string projectName = projRef.Substring(projRef.LastIndexOf("\\") + 1);
			XmlElement xmlProjRef = null;
			foreach (XmlElement element in root.SelectNodes("//ns:ProjectReference", nsMgr))
			{
				if (element.Attributes["Include"] == null)
					continue;

				if (!element.Attributes["Include"].Value.EndsWith(projectName))
					continue;

				xmlProjRef = element;
				//element.ParentNode.RemoveChild(element);
				break;
			}

			if (xmlProjRef == null)
				throw new Exception("Project Reference not found");

			XmlNode insideVS = root.SelectSingleNode("//ns:ItemGroup[Condition=\"'$(BuildingInsideVisualStudio)' == 'true' \"]", nsMgr);
			if (insideVS == null)
			{
				XmlElement group = xml.CreateElement("ItemGroup", "http://schemas.microsoft.com/developer/msbuild/2003");
				XmlAttribute cond = xml.CreateAttribute("Condition");
				cond.Value = "'$(BuildingInsideVisualStudio)' == 'true'";
				group.Attributes.Append(cond);

				group.AppendChild(xmlProjRef);
				root.AppendChild(group);
			}
			else
			{
				insideVS.AppendChild(xmlProjRef);
			}

			XmlElement xmlLibRef = xml.CreateElement("Reference", "http://schemas.microsoft.com/developer/msbuild/2003");
			XmlAttribute nameAtt = xml.CreateAttribute("Include");
			nameAtt.Value = projectName.Substring(0, projectName.LastIndexOf("."));
			XmlElement hint = xml.CreateElement("HintPath", "http://schemas.microsoft.com/developer/msbuild/2003");

			Uri currentUri = new Uri(projectFile);
			Uri refSourceUri = new Uri(libRef);
			Uri diff = currentUri.MakeRelativeUri(refSourceUri);
			string relPath = diff.OriginalString.Replace("/Debug/","/$(Configuration)/").Replace("/Release/","/$(Configuration)/");

			hint.InnerText = relPath.Replace("/","\\");
			XmlElement priv = xml.CreateElement("Private", "http://schemas.microsoft.com/developer/msbuild/2003");
			priv.InnerText = "False";

			xmlLibRef.Attributes.Append(nameAtt);
			xmlLibRef.AppendChild(hint);
			xmlLibRef.AppendChild(priv);

			XmlNode outsideVS = root.SelectSingleNode("//ns:ItemGroup[Condition=\"'$(BuildingInsideVisualStudio)' != 'true' \"]", nsMgr);
			if (outsideVS == null)
			{
				XmlElement group = xml.CreateElement("ItemGroup", "http://schemas.microsoft.com/developer/msbuild/2003");
				XmlAttribute cond = xml.CreateAttribute("Condition");
				cond.Value = "'$(BuildingInsideVisualStudio)' != 'true'";
				group.Attributes.Append(cond);

				group.AppendChild(xmlLibRef);
				root.AppendChild(group);
			}
			else
			{
				outsideVS.AppendChild(xmlLibRef);
			}

			xml.Save(projectFile);
		}

	}
}
