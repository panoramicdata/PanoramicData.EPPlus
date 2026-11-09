/*******************************************************************************
 * You may amend and distribute as you like, but don't remove this header!
 *
 * EPPlus provides server-side generation of Excel 2007/2010 spreadsheets.
 * See https://github.com/JanKallman/EPPlus for details.
 *
 * Copyright (C) 2011  Jan Källman
 *
 * This library is free software; you can redistribute it and/or
 * modify it under the terms of the GNU Lesser General Public
 * License as published by the Free Software Foundation; either
 * version 2.1 of the License, or (at your option) any later version.

 * This library is distributed in the hope that it will be useful,
 * but WITHOUT ANY WARRANTY; without even the implied warranty of
 * MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  
 * See the GNU Lesser General Public License for more details.
 *
 * The GNU Lesser General Public License can be viewed at http://www.opensource.org/licenses/lgpl-license.php
 * If you unfamiliar with this license or have questions about it, here is an http://www.gnu.org/licenses/gpl-faq.html
 *
 * All code and executables are provided "as is" with no warranty either express or implied. 
 * The author accepts no liability for any damage or loss of business that this product may cause.
 *
 * Code change notes:
 * 
 * Author							Change						Date
 * ******************************************************************************
 * Eyal Seagull        Added       		  2012-04-03
 *******************************************************************************/
using OfficeOpenXml.ConditionalFormatting.Contracts;
using System.Xml;

namespace OfficeOpenXml.ConditionalFormatting.Rules;

/// <summary>
/// ExcelConditionalFormattingTwoColorScale
/// </summary>
public class ExcelConditionalFormattingTwoColorScale
  : ExcelConditionalFormattingRule,
  IExcelConditionalFormattingTwoColorScale
{
	/****************************************************************************************/

	#region Private Properties
	/// <summary>
	/// Private Low Value
	/// </summary>
	private ExcelConditionalFormattingColorScaleValue _lowValue;

	/// <summary>
	/// Private High Value
	/// </summary>
	private ExcelConditionalFormattingColorScaleValue _highValue;
	#endregion Private Properties

	/****************************************************************************************/

	#region Constructors
	/// <summary>
	/// 
	/// </summary>
	/// <param name="priority"></param>
	/// <param name="address"></param>
	/// <param name="worksheet"></param>
	/// <param name="itemElementNode"></param>
	/// <param name="namespaceManager"></param>
	internal ExcelConditionalFormattingTwoColorScale(
	  ExcelAddress address,
	  int priority,
	  ExcelWorksheet worksheet,
	  XmlNode itemElementNode = null,
	  XmlNamespaceManager namespaceManager = null)
	  : base(
		eExcelConditionalFormattingRuleType.TwoColorScale,
		address,
		priority,
		worksheet,
		itemElementNode,
		namespaceManager == null ? worksheet.NameSpaceManager : namespaceManager)
	{
		// If the node is not null, parse values out of it instead of clobbering it.
		if (itemElementNode == null)
		{
			// Create the <colorScale> node inside the <cfRule> node
			CreateComplexNode(
			  Node,
			  ExcelConditionalFormattingConstants.Paths.ColorScale);

			LowValue = ExcelConditionalFormattingColorScaleValue.CreateDefaultLow(eExcelConditionalFormattingRuleType.TwoColorScale, address, priority, worksheet, NameSpaceManager);
			HighValue = ExcelConditionalFormattingColorScaleValue.CreateDefaultHigh(eExcelConditionalFormattingRuleType.TwoColorScale, address, priority, worksheet, NameSpaceManager);
		}
	}
	#endregion Constructors

	/****************************************************************************************/

	#region Public Properties
	/// <summary>
	/// Low Value for Two Color Scale Object Value
	/// </summary>
	public ExcelConditionalFormattingColorScaleValue LowValue
	{
		get { return _lowValue; }
		set { _lowValue = value; }
	}

	/// <summary>
	/// High Value for Two Color Scale Object Value
	/// </summary>
	public ExcelConditionalFormattingColorScaleValue HighValue
	{
		get { return _highValue; }
		set { _highValue = value; }
	}
	#endregion Public Properties

	/****************************************************************************************/
}