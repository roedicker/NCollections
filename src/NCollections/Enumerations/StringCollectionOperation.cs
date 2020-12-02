using System;

namespace NCollections.Enumerations
{
  /// <summary>
  /// Enumeration of a single string collcetion operation
  /// </summary>
  [Serializable()]
  public enum StringCollectionOperation
  {
    /// <summary>
    /// Adding a string collection item.
    /// </summary>
    Add = 1,

    /// <summary>
    /// Inserting a string collection item.
    /// </summary>
    Insert = 2,

    /// <summary>
    /// Updating a string collection item.
    /// </summary>
    Update = 3,

    /// <summary>
    /// Removing a string collection item.
    /// </summary>
    Remove = 4
  }
}
