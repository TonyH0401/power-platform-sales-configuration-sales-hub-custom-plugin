#pragma warning disable CS1591
//------------------------------------------------------------------------------
// <auto-generated>
//     This code was generated by a tool.
//
//     Changes to this file may cause incorrect behavior and will be lost if
//     the code is regenerated.
// </auto-generated>
//------------------------------------------------------------------------------

namespace ShPlugins
{
	
	
	/// <summary>
	/// The timeframe that this lead is likely to make a purchase in.
	/// </summary>
	[System.Runtime.Serialization.DataContractAttribute()]
	public enum PurchaseTimeFrame
	{
		
		[System.Runtime.Serialization.EnumMemberAttribute()]
		[OptionSetMetadataAttribute("Immediate", 0)]
		Immediate = 0,
		
		[System.Runtime.Serialization.EnumMemberAttribute()]
		[OptionSetMetadataAttribute("Next Quarter", 2)]
		NextQuarter = 2,
		
		[System.Runtime.Serialization.EnumMemberAttribute()]
		[OptionSetMetadataAttribute("This Quarter", 1)]
		ThisQuarter = 1,
		
		[System.Runtime.Serialization.EnumMemberAttribute()]
		[OptionSetMetadataAttribute("This Year", 3)]
		ThisYear = 3,
		
		[System.Runtime.Serialization.EnumMemberAttribute()]
		[OptionSetMetadataAttribute("Unknown", 4)]
		Unknown = 4,
	}
}
#pragma warning restore CS1591
