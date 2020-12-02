using System;

using NCollections.Enumerations;

namespace NCollections.EventArgs
{
  /// <summary>
  /// Event arguments for a string collection operation.
  /// </summary>
  [Serializable()]
  public class StringCollectionEventArgs
  {
    /// <summary>
    /// Gets the value of the string collection event argument
    /// </summary>
    public string? Value
    {
      get;
    }

    /// <summary>
    /// Gets the operation of the string collection event argument
    /// </summary>
    public StringCollectionOperation Operation
    {
      get;
    }

    /// <summary>
    /// Creates a new instance of the string collection event arguments
    /// </summary>
    /// <param name="value">Value of the event argument</param>
    /// <param name="operation">Operation of the event argument</param>
    public StringCollectionEventArgs(string? value, StringCollectionOperation operation)
    {
      this.Value = value;
      this.Operation = operation;
    }
  }
}
