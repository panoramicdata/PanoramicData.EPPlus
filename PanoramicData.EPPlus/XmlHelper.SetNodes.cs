using System.Xml;

namespace OfficeOpenXml;

public abstract partial class XmlHelper
{
	internal void DeleteAllNode(string path)
	{
		var split = path.Split('/');
		var node = TopNode;
		foreach (var s in split)
		{
			node = node.SelectSingleNode(s, NameSpaceManager);
			if (node != null)
			{
				if (node is XmlAttribute attribute)
				{
					attribute.OwnerElement.Attributes.Remove(attribute);
				}
				else
				{
					node.ParentNode.RemoveChild(node);
				}
			}
			else
			{
				break;
			}
		}
	}
	internal void DeleteNode(string path)
	{
		var node = TopNode.SelectSingleNode(path, NameSpaceManager);
		if (node != null)
		{
			if (node is XmlAttribute att)
			{
				att.OwnerElement.Attributes.Remove(att);
			}
			else
			{
				node.ParentNode.RemoveChild(node);
			}
		}
	}
	internal void DeleteTopNode() => TopNode.ParentNode.RemoveChild(TopNode);
	internal void SetXmlNodeString(string path, string value) => SetXmlNodeString(TopNode, path, value, false, false);
	internal void SetXmlNodeString(string path, string value, bool removeIfBlank) => SetXmlNodeString(TopNode, path, value, removeIfBlank, false);
	internal void SetXmlNodeString(XmlNode node, string path, string value) => SetXmlNodeString(node, path, value, false, false);
	internal void SetXmlNodeString(XmlNode node, string path, string value, bool removeIfBlank) => SetXmlNodeString(node, path, value, removeIfBlank, false);
	internal void SetXmlNodeString(XmlNode node, string path, string value, bool removeIfBlank, bool insertFirst)
	{
		if (node == null)
		{
			return;
		}

		if (value == "" && removeIfBlank)
		{
			DeleteAllNode(path);
		}
		else
		{
			var nameNode = node.SelectSingleNode(path, NameSpaceManager);
			if (nameNode == null)
			{
				CreateNode(path, insertFirst);
				nameNode = node.SelectSingleNode(path, NameSpaceManager);
			}

			nameNode.InnerText = value;
		}
	}
	internal void SetXmlNodeBool(string path, bool value) => SetXmlNodeString(TopNode, path, value ? "1" : "0", false, false);
	internal void SetXmlNodeBool(string path, bool value, bool removeIf)
	{
		if (value == removeIf)
		{
			var node = TopNode.SelectSingleNode(path, NameSpaceManager);
			if (node != null)
			{
				if (node is XmlAttribute attribute)
				{
					var elem = attribute.OwnerElement;
					elem.ParentNode.RemoveChild(elem);
				}
				else
				{
					TopNode.RemoveChild(node);
				}
			}
		}
		else
		{
			SetXmlNodeString(TopNode, path, value ? "1" : "0", false, false);
		}
	}
	/// <summary>
	/// Insert the new node before any of the nodes in the comma separated list
	/// </summary>
	/// <param name="parentNode">Parent node</param>
	/// <param name="beforeNodes">comma separated list containing nodes to insert after. Left to right order</param>
	/// <param name="newNode">The new node to be inserted</param>
	internal void InserAfter(XmlNode parentNode, string beforeNodes, XmlNode newNode)
	{
		var nodePaths = beforeNodes.Split(',');

		foreach (var nodePath in nodePaths)
		{
			var node = parentNode.SelectSingleNode(nodePath, NameSpaceManager);
			if (node != null)
			{
				parentNode.InsertAfter(newNode, node);
				return;
			}
		}

		parentNode.InsertAfter(newNode, null);
	}
}
