using System.Xml;

namespace OfficeOpenXml;

public abstract partial class XmlHelper
{
	internal XmlNode CreateNode(string path) => path == "" ? TopNode : CreateNode(path, false);

	/// <summary>
	/// Create the node path. Nodesa are inserted according to the Schema node oreder
	/// </summary>
	/// <param name="path">The path to be created</param>
	/// <param name="insertFirst">Insert as first child</param>
	/// <param name="addNew">Always add a new item at the last level.</param>
	/// <returns></returns>
	internal XmlNode CreateNode(string path, bool insertFirst, bool addNew = false)
	{
		var node = TopNode;
		XmlNode prependNode = null;
		var first = insertFirst;
		if (path.StartsWith('/'))
		{
			path = path[1..];
		}

		var subPaths = path.Split('/');
		for (var i = 0; i < subPaths.Length; i++)
		{
			var subPath = subPaths[i];
			var subNode = node.SelectSingleNode(subPath, NameSpaceManager);
			if (subNode == null || (i == subPath.Length - 1 && addNew))
			{
				if (SchemaNodeOrder != null && subPath[0] != '@')
				{
					first = false;
					prependNode = GetPrependNode(subPath, node);
				}

				SplitNodeName(subPath, out var nodePrefix, out var nodeName, out var nameSpaceURI);
				if (subPath.StartsWith('@'))
				{
					var addedAtt = node.OwnerDocument.CreateAttribute(subPath[1..], nameSpaceURI);  //nameSpaceURI
					node.Attributes.Append(addedAtt);
				}
				else
				{
					subNode = CreateNodeElement(node, nodePrefix, nodeName, nameSpaceURI);
					if (prependNode != null)
					{
						node.InsertBefore(subNode, prependNode);
						prependNode = null;
					}
					else if (first)
					{
						node.PrependChild(subNode);
					}
					else
					{
						node.AppendChild(subNode);
					}
				}
			}

			node = subNode;
		}

		return node;
	}

	private void SplitNodeName(string subPath, out string nodePrefix, out string nodeName, out string nameSpaceURI)
	{
		var nameSplit = subPath.Split(':');
		if (nameSplit.Length > 1)
		{
			nodePrefix = nameSplit[0];
			if (nodePrefix[0] == '@')
			{
				nodePrefix = nodePrefix[1..];
			}

			nameSpaceURI = NameSpaceManager.LookupNamespace(nodePrefix);
			nodeName = nameSplit[1];
		}
		else
		{
			nodePrefix = "";
			nameSpaceURI = "";
			nodeName = nameSplit[0];
		}
	}

	private static XmlNode CreateNodeElement(XmlNode node, string nodePrefix, string nodeName, string nameSpaceURI)
	{
		if (nodePrefix == "")
		{
			return node.OwnerDocument.CreateElement(nodeName, nameSpaceURI);
		}

		if (nodePrefix == "" || (node.OwnerDocument != null && node.OwnerDocument.DocumentElement != null && node.OwnerDocument.DocumentElement.NamespaceURI == nameSpaceURI &&
				node.OwnerDocument.DocumentElement.Prefix == ""))
		{
			return node.OwnerDocument.CreateElement(nodeName, nameSpaceURI);
		}

		return node.OwnerDocument.CreateElement(nodePrefix, nodeName, nameSpaceURI);
	}
	/// <summary>
	/// return Prepend node
	/// </summary>
	/// <param name="nodeName">name of the node to check</param>
	/// <param name="node">Top node to check children</param>
	/// <returns></returns>
	private XmlNode GetPrependNode(string nodeName, XmlNode node)
	{
		var pos = GetNodePos(nodeName);
		if (pos < 0)
		{
			return null;
		}

		XmlNode prependNode = null;
		foreach (XmlNode childNode in node.ChildNodes)
		{
			var childPos = GetNodePos(childNode.Name);
			if (childPos > -1)  //Found?
			{
				if (childPos > pos) //Position is before
				{
					prependNode = childNode;
					break;
				}
			}
		}

		return prependNode;
	}
	private int GetNodePos(string nodeName)
	{
		var ix = nodeName.IndexOf(':');
		if (ix > 0)
		{
			nodeName = nodeName[(ix + 1)..];
		}

		for (var i = 0; i < _schemaNodeOrder.Length; i++)
		{
			if (nodeName == _schemaNodeOrder[i])
			{
				return i;
			}
		}

		return -1;
	}
}
