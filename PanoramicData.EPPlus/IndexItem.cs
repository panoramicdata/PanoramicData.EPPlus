using System.Diagnostics.CodeAnalysis;

[SuppressMessage("Minor Code Smell", "S1210:Equality and comparison operators should be overloaded when implementing IComparable", Justification = "Internal ordering key for Array.BinarySearch; overriding equality would change reference semantics.")]
internal class IndexBase : IComparable<IndexBase>
{
	internal short Index;
	public int CompareTo(IndexBase other) =>
		Index - other.Index;
}
// to compress memory size, use struct
[SuppressMessage("Minor Code Smell", "S1210:Equality and comparison operators should be overloaded when implementing IComparable", Justification = "Internal ordering key for Array.BinarySearch only.")]
[SuppressMessage("Minor Code Smell", "S3898:Value types should implement IEquatable<T>", Justification = "Internal ordering key for Array.BinarySearch only; never compared for equality.")]
internal struct IndexItem : IComparable<IndexItem>
{
	internal int IndexPointer
	{
		get;
		set;
	}
	internal short Index;
	public int CompareTo(IndexItem other) => Index - other.Index;
}
