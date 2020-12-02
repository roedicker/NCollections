using System;
using System.Collections;
using System.Collections.Generic;
using System.Globalization;

using NCollections.Enumerations;

namespace NCollections.Comparers
{
  /// <summary>
  /// Comparer for string related collections. <see cref="SortDirection" />
  /// class will be used to handle sort directions.
  /// </summary>
  [Serializable()]
  internal class StringCollectionComparer : IComparer, IEqualityComparer<string>
  {
    /// <summary>
    /// Gets or sets the sort direction for the string collection.
    /// </summary>
    /// <value>Sort direction - either ascending or descending.</value>
    /// <returns>Current direction for sorting the collection.</returns>
    internal SortDirection SortDirection
    {
      get;
      set;
    }

    /// <summary>
    /// Gets or sets the indicator whether case is ignored for comparing two string items or not.
    /// </summary>
    /// <value>Indicator whether case is ignored for comparing two string items or not.</value>
    /// <returns>Current indicator whether case is ignored for comparing two string items or not.</returns>
    /// <remarks>The default value is <strong>false</strong>.</remarks>
    public bool IgnoreCase
    {
      get;
      set;
    }

    /// <summary>
    /// Initializes a new instance of the sort comparer for string collections.
    /// </summary>
    public StringCollectionComparer()
    {
      this.SortDirection = SortDirection.Ascending;
      this.IgnoreCase = false;
    }

    /// <summary>
    /// Initializes a new instance of the sort comparer for string collections.
    /// </summary>
    public StringCollectionComparer(bool ignoreCase) : this()
    {
      this.IgnoreCase = ignoreCase;
    }

    /// <summary>
    /// Method for comparing two single items of the string collection.
    /// </summary>
    /// <param name="x">First item to compare.</param>
    /// <param name="y">Second item to compare.</param>
    /// <returns>1 if first item is greater that the second, -1 if the second item is greater than the first one and 0 if both items have identifally values</returns>
    /// <remarks>The result is equal to a call of the <see cref="String.Compare(String, String, bool)" /> function.</remarks>
    public int Compare(object x, object y)
    {
      int Result;

      if (x == null)
      {
        throw new ArgumentNullException(nameof(x));
      }

      if (x as string == null)
      {
        throw new ArgumentException(Resources.CompareFailedArgumentMustBeOfTypeStringErrorMessage, nameof(x));
      }

      if (y == null)
      {
        throw new ArgumentNullException(nameof(y));
      }

      if (y as string == null)
      {
        throw new ArgumentException(Resources.CompareFailedArgumentMustBeOfTypeStringErrorMessage, nameof(y));
      }

      string sX = (string)x;
      string sY = (string)y;

      Result = String.Compare(sX, sY, this.IgnoreCase ? StringComparison.OrdinalIgnoreCase : StringComparison.Ordinal);

      if (this.SortDirection == SortDirection.Descending)
      {
        Result *= -1;
      }

      return Result;
    }

    /// <summary>
    /// Checks the equality of two strings.
    /// </summary>
    /// <param name="x">First string to compare.</param>
    /// <param name="y">Second string to compare.</param>
    /// <returns><strong>true</strong> if both strings are equal, otherwise <strong>false</strong>.</returns>
    /// <remarks>Result may depend on the <see cref="IgnoreCase" /> property of the comparer.</remarks>
    public bool Equals(string x, string y)
    {
      return String.Compare(x, y, this.IgnoreCase ? StringComparison.OrdinalIgnoreCase : StringComparison.Ordinal) == 0;
    }

    /// <summary>
    /// Checks the equality of this instance
    /// </summary>
    /// <param name="obj">Other instance to compare for equality</param>
    /// <returns><c>True</c> if other instance is equal to this instance, otherwise <c>False</c></returns>
    public override bool Equals(object? obj)
    {
      return obj is StringCollectionComparer comparer &&
             this.SortDirection == comparer.SortDirection &&
             this.IgnoreCase == comparer.IgnoreCase;
    }

    /// <summary>
    /// Gets the Hashcode of a single string item.
    /// </summary>
    /// <param name="s">String item to get hashcode for.</param>
    /// <returns>Hashcode for string item.</returns>
    public int GetHashCode(string s)
    {
      int Result = 0;

      if (!String.IsNullOrWhiteSpace(s))
      {
        if (this.IgnoreCase)
        {
          s = s.ToUpper(CultureInfo.InvariantCulture);
        }

        Result = s.GetHashCode();
      }

      return Result;
    }
    /// <summary>
    /// Gets the hascode for this instance
    /// </summary>
    /// <returns>Hashcode for this instance</returns>
    public override int GetHashCode()
    {
      int hashCode = -631605227;
      hashCode = hashCode * -1521134295 + this.SortDirection.GetHashCode();
      hashCode = hashCode * -1521134295 + this.IgnoreCase.GetHashCode();
      return hashCode;
    }
  }
}
