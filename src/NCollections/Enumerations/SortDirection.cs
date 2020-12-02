using System;

namespace NCollections.Enumerations
{
  /// <summary>
  /// Defines sort direction for collections.
  /// </summary>
  [Serializable]
  public enum SortDirection
  {
    /// <summary>
    /// No sort has to be performed.
    /// </summary>
    None = 0,

    /// <summary>
    /// Sort ascending.
    /// </summary>
    Ascending = 1,

    /// <summary>
    /// Sort descending.
    /// </summary>
    Descending = 2
  }
}
