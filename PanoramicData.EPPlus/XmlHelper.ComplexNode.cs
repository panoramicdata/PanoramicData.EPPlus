using System.Diagnostics.CodeAnalysis;
using System.Xml;

namespace OfficeOpenXml;

public abstract partial class XmlHelper
{
	/// <summary>
	/// Create a complex node. Insert the node according to SchemaOrder
	/// using the TopNode as the parent
	/// </summary>
	/// <param name="path"></param>
	/// <returns></returns>
	internal XmlNode CreateComplexNode(
		string path) => CreateComplexNode(
			TopNode,
			path,
			eNodeInsertOrder.SchemaOrder,
			null);

	/// <summary>
	/// Create a complex node. Insert the node according to the <paramref name="path"/>
	/// using the <paramref name="topNode"/> as the parent
	/// </summary>
	/// <param name="topNode"></param>
	/// <param name="path"></param>
	/// <returns></returns>
	internal XmlNode CreateComplexNode(
		XmlNode topNode,
		string path) => CreateComplexNode(
			topNode,
			path,
			eNodeInsertOrder.SchemaOrder,
			null);

	/// <summary>
	/// Creates complex XML nodes
	/// </summary>
	/// <remarks>
	/// 1. "d:conditionalFormatting"
	///		1.1. Creates/find the first "conditionalFormatting" node
	/// 
	/// 2. "d:conditionalFormatting/@sqref"
	///		2.1. Creates/find the first "conditionalFormatting" node
	///		2.2. Creates (if not exists) the @sqref attribute
	///
	/// 3. "d:conditionalFormatting/@id='7'/@sqref='A9:B99'"
	///		3.1. Creates/find the first "conditionalFormatting" node
	///		3.2. Creates/update its @id attribute to "7"
	///		3.3. Creates/update its @sqref attribute to "A9:B99"
	///
	/// 4. "d:conditionalFormatting[@id='7']/@sqref='X1:X5'"
	///		4.1. Creates/find the first "conditionalFormatting" node with @id=7
	///		4.2. Creates/update its @sqref attribute to "X1:X5"
	///	
	/// 5. "d:conditionalFormatting[@id='7']/@id='8'/@sqref='X1:X5'/d:cfRule/@id='AB'"
	///		5.1. Creates/find the first "conditionalFormatting" node with @id=7
	///		5.2. Set its @id attribute to "8"
	///		5.2. Creates/update its @sqref attribute and set it to "X1:X5"
	///		5.3. Creates/find the first "cfRule" node (inside the node)
	///		5.4. Creates/update its @id attribute to "AB"
	///	
	/// 6. "d:cfRule/@id=''"
	///		6.1. Creates/find the first "cfRule" node
	///		6.1. Remove the @id attribute
	///	</remarks>
	/// <param name="topNode"></param>
	/// <param name="path"></param>
	/// <param name="nodeInsertOrder"></param>
	/// <param name="referenceNode"></param>
	/// <returns>The last node creates/found</returns>
	internal XmlNode CreateComplexNode(
		XmlNode topNode,
		string path,
		eNodeInsertOrder nodeInsertOrder,
		XmlNode? referenceNode)
	{
		// Path is obligatory
		if ((path == null) || (path == string.Empty))
		{
			return topNode;
		}

		var node = topNode;

		//TODO: BUG: when the "path" contains "/" in an attribute value, it gives an error.

		// Separate the XPath to Nodes and Attributes
		foreach (var subPath in path.Split('/'))
		{
			// The subPath can be any one of those:
			// nodeName
			// x:nodeName
			// nodeName[find criteria]
			// x:nodeName[find criteria]
			// @attribute
			// @attribute='attribute value'

			// Check if the subPath has at least one character
			if (subPath.Length > 0)
			{
				// Check if the subPath is an attribute (with or without value)
				if (subPath.StartsWith('@'))
				{
					SetComplexNodeAttribute(node, subPath);
				}
				else
				{
					// Look for the node (with or without filter criteria)
					var subNode = node.SelectSingleNode(subPath, NameSpaceManager);

					// Check if the node does not exists
					if (subNode == null)
					{
						subNode = CreateComplexChildElement(node, subPath, out var nodeName);
						InsertComplexChild(node, subNode, nodeName, ref nodeInsertOrder, ref referenceNode);
					}

					// Make the newly created node the top node when the rest of the path
					// is being evaluated. So newly created nodes will be the children of the
					// one we just created.
					node = subNode;
				}
			}
		}

		// Return the last created/found node
		return node;
	}

	private static void SetComplexNodeAttribute(XmlNode node, string subPath)
	{
		// @attribute										--> Create attribute
		// @attribute=''								--> Remove attribute
		// @attribute='attribute value' --> Create attribute + update value
		var attributeSplit = subPath.Split('=');
		var attributeName = attributeSplit[0][1..];
		string? attributeValue = null;   // Null means no attribute value

		// Check if we have an attribute value to set
		if (attributeSplit.Length > 1)
		{
			// Remove the ' or " from the attribute value
			attributeValue = attributeSplit[1].Replace("'", "").Replace("\"", "");
		}

		// Get the attribute (if exists)
		var attribute = (XmlAttribute)(node.Attributes.GetNamedItem(attributeName));

		// Remove the attribute if value is empty (not null)
		if (attributeValue == string.Empty)
		{
			// Only if the attribute exists
			if (attribute != null)
			{
				node.Attributes.Remove(attribute);
			}
		}
		else
		{
			// Create the attribute if does not exists
			if (attribute == null)
			{
				// Create the attribute
				attribute = node.OwnerDocument.CreateAttribute(
					attributeName);

				// Add it to the current node
				node.Attributes.Append(attribute);
			}

			// Update the attribute value
			if (attributeValue != null)
			{
				node.Attributes[attributeName].Value = attributeValue;
			}
		}
	}

	[SuppressMessage("Minor Bug", "S2692:IndexOf checks should not be for positive numbers", Justification = "A criteria bracket at index 0 would leave no node name, so only a later bracket is stripped.")]
	private XmlNode CreateComplexChildElement(XmlNode node, string subPath, out string nodeName)
	{
		// nodeName
		// x:nodeName
		// nodeName[find criteria]
		// x:nodeName[find criteria]
		string nodePrefix;
		string nameSpaceURI;
		var nameSplit = subPath.Split(':');

		// Check if the name has a prefix like "d:nodeName"
		if (nameSplit.Length > 1)
		{
			nodePrefix = nameSplit[0];
			nameSpaceURI = NameSpaceManager.LookupNamespace(nodePrefix);
			nodeName = nameSplit[1];
		}
		else
		{
			nodePrefix = string.Empty;
			nameSpaceURI = string.Empty;
			nodeName = nameSplit[0];
		}

		// Check if we have a criteria part in the node name
		if (nodeName.IndexOf('[') > 0)
		{
			// remove the criteria from the node name
			nodeName = nodeName[..nodeName.IndexOf('[')];
		}

		if (nodePrefix == string.Empty)
		{
			return node.OwnerDocument.CreateElement(nodeName, nameSpaceURI);
		}

		if (node.OwnerDocument != null
			&& node.OwnerDocument.DocumentElement != null
			&& node.OwnerDocument.DocumentElement.NamespaceURI == nameSpaceURI
			&& node.OwnerDocument.DocumentElement.Prefix == string.Empty)
		{
			return node.OwnerDocument.CreateElement(
				nodeName,
				nameSpaceURI);
		}

		return node.OwnerDocument.CreateElement(
			nodePrefix,
			nodeName,
			nameSpaceURI);
	}

	// The insert order and reference node carry over to the following path segments.
	private void InsertComplexChild(XmlNode node, XmlNode subNode, string nodeName, ref eNodeInsertOrder nodeInsertOrder, ref XmlNode? referenceNode)
	{
		// Check if we need to use the "SchemaOrder"
		if (nodeInsertOrder == eNodeInsertOrder.SchemaOrder)
		{
			// Check if the Schema Order List is empty
			if (SchemaNodeOrder == null || SchemaNodeOrder.Length == 0)
			{
				// Use the "Insert Last" option when Schema Order List is empty
				nodeInsertOrder = eNodeInsertOrder.Last;
			}
			else
			{
				// Find the prepend node in order to insert
				referenceNode = GetPrependNode(nodeName, node);

				if (referenceNode != null)
				{
					nodeInsertOrder = eNodeInsertOrder.Before;
				}
				else
				{
					nodeInsertOrder = eNodeInsertOrder.Last;
				}
			}
		}

		switch (nodeInsertOrder)
		{
			case eNodeInsertOrder.After:
				node.InsertAfter(subNode, referenceNode);
				referenceNode = null;
				break;

			case eNodeInsertOrder.Before:
				node.InsertBefore(subNode, referenceNode);
				referenceNode = null;
				break;

			case eNodeInsertOrder.First:
				node.PrependChild(subNode);
				break;

			case eNodeInsertOrder.Last:
				node.AppendChild(subNode);
				break;
		}
	}
}