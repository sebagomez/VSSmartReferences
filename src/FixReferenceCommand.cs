//------------------------------------------------------------------------------
// <copyright file="FixReferenceCommand.cs" company="SebaGomez">
//     Copyright (c) SebaGomez.  All rights reserved.
// </copyright>
//------------------------------------------------------------------------------

using System;
using System.ComponentModel.Design;
using System.Xml;
using EnvDTE;
using EnvDTE80;
using Microsoft.VisualStudio;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using VSSmartReferences.Helpers;

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

		private void FixReferenceCallback(object sender, EventArgs e)
		{
			var dte = (DTE2)ServiceProvider.GetService(typeof(DTE));
			foreach (object item in CommandUtils.GetSelectedReferences(dte))
			{
				var d = (dynamic)item;
				if (d.SourceProject == null)
				{
					Writeline($"{d.Name} is not a source project");
					continue;
				}

				string libPath = d.Path;
				string sourceProject = d.SourceProject.FullName;
				string currentProject = d.ContainingProject.FullName;


				FixReference(currentProject, libPath, sourceProject);
			}

		}

		const string CSPROJ_NAMESPACE = "http://schemas.microsoft.com/developer/msbuild/2003";
		const string CSPROJ_NAMESPACE_PREFIX = "ns";

		void FixReference(string projectFile, string libRef, string projRef)
		{
			XmlDocument xml = GetCSProjDocument(projectFile, out XmlNamespaceManager nsMgr);
			string projectName = projRef.Substring(projRef.LastIndexOf("\\") + 1);
			XmlElement xmlProjRef = GetProjectReference(xml, nsMgr, projectName);

			if (xmlProjRef == null)
			{
				Writeline($"Project reference for {projectName} not found");
				return;
			}

			AddInsideVSRef(xml, nsMgr, xmlProjRef);
			AddOutsideVSRef(xml, nsMgr, projectName, projectFile, libRef);

			xml.Save(projectFile);
		}

		private void AddOutsideVSRef(XmlDocument xml, XmlNamespaceManager nsMgr, string projectName, string projectFile, string libRef)
		{
			XmlElement xmlLibRef = xml.CreateElement(CommandUtils.REFERENCE, CSPROJ_NAMESPACE);
			XmlAttribute nameAtt = xml.CreateAttribute(CommandUtils.INCLUDE);
			nameAtt.Value = projectName.Substring(0, projectName.LastIndexOf("."));
			XmlElement hint = xml.CreateElement(CommandUtils.HINTPATH, CSPROJ_NAMESPACE);

			Uri currentUri = new Uri(projectFile);
			Uri refSourceUri = new Uri(libRef);
			Uri diff = currentUri.MakeRelativeUri(refSourceUri);
			string relPath = diff.OriginalString.Replace("/Debug/", "/$(Configuration)/").Replace("/Release/", "/$(Configuration)/");

			hint.InnerText = relPath.Replace("/", "\\");
			XmlElement priv = xml.CreateElement(CommandUtils.PRIVATE, CSPROJ_NAMESPACE);
			priv.InnerText = bool.FalseString;

			xmlLibRef.Attributes.Append(nameAtt);
			xmlLibRef.AppendChild(hint);
			xmlLibRef.AppendChild(priv);

			XmlNode outsideVS = xml.DocumentElement.SelectSingleNode($"//{CSPROJ_NAMESPACE_PREFIX}:ItemGroup[@Condition=\"{CommandUtils.BUILD_OUTSIDE_VS}\"]", nsMgr);
			if (outsideVS == null)
			{
				XmlElement group = xml.CreateElement(CommandUtils.ITEM_GORUP, CSPROJ_NAMESPACE);
				XmlAttribute cond = xml.CreateAttribute(CommandUtils.CONDITION);
				cond.Value = CommandUtils.BUILD_OUTSIDE_VS;
				group.Attributes.Append(cond);

				group.AppendChild(xmlLibRef);
				xml.DocumentElement.AppendChild(group);
			}
			else
			{
				outsideVS.AppendChild(xmlLibRef);
			}
		}

		private void AddInsideVSRef(XmlDocument xml, XmlNamespaceManager nsMgr, XmlElement xmlProjRef)
		{
			SetNoCopyLocal(xmlProjRef, nsMgr);

			if (xmlProjRef.ParentNode.Attributes[CommandUtils.CONDITION] != null)
			{
				if (xmlProjRef.ParentNode.Attributes[CommandUtils.CONDITION].Value == CommandUtils.BUILD_INSIDE_VS)
					return;
			}

			XmlNode insideVS = xml.DocumentElement.SelectSingleNode($"//{CSPROJ_NAMESPACE_PREFIX}:ItemGroup[@Condition=\"{CommandUtils.BUILD_INSIDE_VS}\"]", nsMgr);
			if (insideVS == null)
			{
				XmlElement group = xml.CreateElement(CommandUtils.ITEM_GORUP, CSPROJ_NAMESPACE);
				XmlAttribute cond = xml.CreateAttribute(CommandUtils.CONDITION);
				cond.Value = CommandUtils.BUILD_INSIDE_VS;
				group.Attributes.Append(cond);

				group.AppendChild(xmlProjRef);
				xml.DocumentElement.AppendChild(group);
			}
			else
			{
				insideVS.AppendChild(xmlProjRef);
			}
		}

		private void SetNoCopyLocal(XmlElement xmlProjRef, XmlNamespaceManager nsMgr)
		{
			XmlNode priv = xmlProjRef.SelectSingleNode($"{CSPROJ_NAMESPACE_PREFIX}:{CommandUtils.PRIVATE}", nsMgr);
			if (priv == null)
				priv = xmlProjRef.OwnerDocument.CreateElement(CommandUtils.PRIVATE, CSPROJ_NAMESPACE);
			
			priv.InnerText = bool.FalseString;
			if (priv.ParentNode == null)
				xmlProjRef.AppendChild(priv);
		}

		private XmlElement GetProjectReference(XmlDocument xml, XmlNamespaceManager nsMgr, string projectName)
		{
			foreach (XmlElement element in xml.DocumentElement.SelectNodes($"//{CSPROJ_NAMESPACE_PREFIX}:{CommandUtils.PROJECT_REFERENCE}", nsMgr))
			{
				if (element.Attributes[CommandUtils.INCLUDE] == null)
					continue;

				if (!element.Attributes[CommandUtils.INCLUDE].Value.EndsWith(projectName))
					continue;

				return element;
			}

			return null;
		}

		XmlDocument GetCSProjDocument(string currentProject, out XmlNamespaceManager nsMgr)
		{
			XmlDocument xml = new XmlDocument();
			xml.Load(currentProject);
			nsMgr = new XmlNamespaceManager(xml.NameTable);
			nsMgr.AddNamespace(CSPROJ_NAMESPACE_PREFIX, CSPROJ_NAMESPACE);

			return xml;
		}

		void Writeline(string text)
		{
			ThreadHelper.ThrowIfNotOnUIThread();
			IVsOutputWindow outWindow = Package.GetGlobalService(typeof(SVsOutputWindow)) as IVsOutputWindow;

			Guid customGuid = new Guid("C56A383C-7C98-4C13-9D72-81C9F9817BD9");
			string customTitle = "Smart References";
			outWindow.CreatePane(ref customGuid, customTitle, 1, 1);

			IVsOutputWindowPane customPane;
			outWindow.GetPane(ref customGuid, out customPane);

			customPane.OutputString($"{text}\r\n");
			customPane.Activate(); // Brings this pane into view
		}
	}
}
