using System;
using System.Collections;
using System.Collections.Generic;
using System.Globalization;
using System.Text;

using NCollections.Comparers;
using NCollections.Enumerations;
using NCollections.EventArgs;

namespace NCollections
{
  /// <summary>
  /// Defines the event handler for changed items
  /// </summary>
  /// <param name="sender">Sender of the event</param>
  /// <param name="e">Arguments of the event</param>
  public delegate void ItemsEventHandler(object sender, StringCollectionEventArgs e);

  /// <summary>
  /// Represents a sortable collection of strings.
  /// </summary>
  [Serializable()]
  public class StringCollection : CollectionBase, IDisposable
  {
    /// <summary>
    /// Event handler raised for each changed item
    /// </summary>
    public event ItemsEventHandler? ItemsChanged;

    /// <summary>
    /// Gets the sort comparer for sorting the string collection.
    /// </summary>
    /// <returns>Sort comparer for sorting the string items.</returns>
    internal StringCollectionComparer SortComparer
    {
      get;
      private set;
    }

    /// <summary>
    /// Gets or sets the element at the specified index.
    /// </summary>
    /// <param name="index">Ordinal number of string entry within the
    /// collection.</param>
    /// <value>Value of stored string.</value>
    public string this[int index]
    {
      get
      {
        return (string)base.List[index];
      }

      set
      {
        base.List[index] = value;
        RaiseItemsChangedEvent(value, StringCollectionOperation.Update);
      }
    }

    /// <summary>
    /// Gets or sets the element at the specified name.
    /// </summary>
    /// <param name="key">Name of string entry within the
    /// collection.</param>
    /// <value>Value of stored string.</value>
    public string this[string key]
    {
      get
      {
        return (string)base.List[base.List.IndexOf(key)];
      }

      set
      {
        base.List[base.List.IndexOf(key)] = value;

        RaiseItemsChangedEvent(value, StringCollectionOperation.Update);
      }
    }

    /// <summary>
    /// Gets or sets the indicator whether empty entries have to be ignored (e.g. via Add() or Insert()) or not.
    /// </summary>
    public bool IgnoreEmpty
    {
      get;
      set;
    }

    /// <summary>
    /// Initializes a new instance of a string collection.
    /// </summary>
    public StringCollection()
    {
      _Disposed = false;

      this.IgnoreEmpty = true;
      this.SortComparer = new StringCollectionComparer();
      this.ItemsChanged = null;
    }

    /// <summary>
    /// Initializes a new instance of a string collection with an initial string value.
    /// </summary>
    /// <param name="value">Initial string value</param>
    public StringCollection(string value) : this()
    {
      AddDistinct(value);
    }

    /// <summary>
    /// Initializes a new instance of a string collection with an initial distinct array of string values.
    /// </summary>
    /// <param name="values">Initial string array values.</param>
    public StringCollection(string[] values) : this()
    {
      AddRangeDistinct(values);
    }

    /// <summary>
    /// Initializes a new instance of a string collection with an initial distinct enumeration of string values.
    /// </summary>
    /// <param name="values">Initial enumerated string values.</param>
    public StringCollection(IEnumerable<string> values) : this()
    {
      AddRangeDistinct(values);
    }

    /// <summary>
    /// Initializes a new instance of a string collection with an initial distinct collection of string values.
    /// </summary>
    /// <param name="values">Initial string collection values.</param>
    public StringCollection(StringCollection values) : this()
    {
      AddRangeDistinct(values);
    }

    /// <summary>
    /// Class specific implementation of the IDisposable interface.
    /// </summary>
    /// <param name="disposing">Indicates if a dispose was initiated by the class itself or by system.</param>
    protected virtual void Dispose(bool disposing)
    {
      if (!_Disposed)
      {
        if (disposing)
        {
          // not used
        }
      }

      _Disposed = true;
    }

    /// <summary>
    /// Implementation if the IDisposable interface.
    /// </summary>
    /// <remarks>This code added by Visual Basic to correctly implement the disposable pattern.</remarks>
    public void Dispose()
    {
      Dispose(true);
      GC.SuppressFinalize(this);
    }

    /// <summary>
    /// Adds a new string to the end of the string collection.
    /// </summary>
    /// <param name="value">String to add to the list</param>
    /// <returns>Ordinal number of the latest added string.</returns>
    public int Add(string value)
    {
      if (String.IsNullOrEmpty(value))
      {
        if (this.IgnoreEmpty || value == null)
        {
          return -1;
        }
        else
        {
          RaiseItemsChangedEvent(value, StringCollectionOperation.Add);

          return base.List.Add(value);
        }
      }
      else
      {
        RaiseItemsChangedEvent(value, StringCollectionOperation.Add);

        return base.List.Add(value);
      }
    }

    /// <summary>
    /// Adds a new formatted string to the end of the string collection.
    /// </summary>
    /// <param name="format">A composite format string.</param>
    /// <param name="args">An System.Object array containing zero or more objects to format.</param>
    /// <returns>Ordinal number of the latest added string.</returns>
    public int AddFormat(string format, params object[] args)
    {
      return Add(String.Format(CultureInfo.InvariantCulture, format, args));
    }

    /// <summary>
    /// Adds a new string to the end of the string collection. If that string already is in this collection it will not be added again.
    /// </summary>
    /// <param name="value">String to add to the list.</param>
    /// <returns>Ordinal number of the latest added string. If the item already exists the ordinal number of the existing item will be returned.</returns>
    public int AddDistinct(string value)
    {
      int Result = base.List.IndexOf(value);

      // check if string already exists in the list, otherwise add it
      if (Result == -1)
      {
        Result = Add(value);
      }

      return Result;
    }

    /// <summary>
    /// Adds a new formatted string to the end of the string collection. If that string already is in this collection it will not be added again.
    /// </summary>
    /// <param name="format">A composite format string.</param>
    /// <param name="args">An System.Object array containing zero or more objects to format.</param>
    /// <returns>Ordinal number of the latest added string.If the item already exists the ordinal number of the existing item will be returned.</returns>
    public int AddDistinctFormat(string format, params object[] args)
    {
      return AddDistinct(String.Format(CultureInfo.InvariantCulture, format, args));
    }

    /// <summary>
    /// Copies the elements of a string array to the end of the StringCollection. 
    /// </summary>
    /// <param name="values">String array to add.</param>
    public virtual void AddRange(string[] values)
    {
      foreach (string sItem in values)
      {
        Add(sItem);
      }
    }

    /// <summary>
    /// Copies the elements of a StringCollection to the end of this StringCollection. 
    /// </summary>
    /// <param name="values">String collection to add.</param>
    public virtual void AddRange(StringCollection values)
    {
      foreach (string sItem in values)
      {
        Add(sItem);
      }
    }

    /// <summary>
    /// Copies the elements of a enumerable to the end of this StringCollection. 
    /// </summary>
    /// <param name="values">enumerable to add.</param>
    public virtual void AddRange(IEnumerable<string> values)
    {
      foreach (string sItem in values)
      {
        Add(sItem);
      }
    }

    /// <summary>
    /// Copies the elements of a key collection to the end of this StringCollection. 
    /// </summary>
    /// <param name="values">Generic string list to add.</param>
    public virtual void AddRange(System.Collections.Specialized.NameObjectCollectionBase.KeysCollection values)
    {
      foreach (string sItem in values)
      {
        Add(sItem);
      }
    }

    /// <summary>
    /// Copies the elements of a string array to the end of the StringCollection. Only non-existing elements will be added.
    /// </summary>
    /// <param name="values">String array to add.</param>
    public virtual void AddRangeDistinct(string[] values)
    {
      foreach (string sItem in values)
      {
        AddDistinct(sItem);
      }
    }

    /// <summary>
    /// Copies the elements of a StringCollection to the end of the StringCollection. Only non-existing elements will be added.
    /// </summary>
    /// <param name="values">String collection to add.</param>
    public virtual void AddRangeDistinct(StringCollection values)
    {
      foreach (string sItem in values)
      {
        AddDistinct(sItem);
      }
    }

    /// <summary>
    /// Copies the elements of a enumerable to the end of the StringCollection. Only non-existing elements will be added.
    /// </summary>
    /// <param name="values">Enumerable to add.</param>
    public virtual void AddRangeDistinct(IEnumerable<string> values)
    {
      foreach (string sItem in values)
      {
        AddDistinct(sItem);
      }
    }

    /// <summary>
    /// Copies the elements of a key collection to the end of this StringCollection. Only non-existing elements will be added.
    /// </summary>
    /// <param name="values">Generic string list to add.</param>
    public virtual void AddRangeDistinct(System.Collections.Specialized.NameObjectCollectionBase.KeysCollection values)
    {
      foreach (string sItem in values)
      {
        AddDistinct(sItem);
      }
    }

    /// <summary>
    /// Assigns an array of strings to this StringCollection. All existing data will be overwritten.
    /// </summary>
    /// <param name="values">Array of strings.</param>
    public virtual void Assign(string[] values)
    {
      Clear();
      AddRange(values);
    }

    /// <summary>
    /// Assigns a StringCollection to this StringCollection. All existing data will be overwritten.
    /// </summary>
    /// <param name="values">StringCollection.</param>
    public virtual void Assign(StringCollection values)
    {
      Clear();
      AddRange(values);
    }

    /// <summary>
    /// Assigns an enumerable to this StringCollection. All existing data will be overwritten.
    /// </summary>
    /// <param name="values">Enumerable of strings to assign.</param>
    public virtual void Assign(IEnumerable<string> values)
    {
      Clear();
      AddRange(values);
    }

    /// <summary>
    /// Assigns a key string collection to this StringCollection. All existing data will be overwritten.
    /// </summary>
    /// <param name="values">Key string collection.</param>
    public virtual void Assign(System.Collections.Specialized.NameObjectCollectionBase.KeysCollection values)
    {
      Clear();
      AddRange(values);
    }

    /// <summary>
    /// Pads the string collection to ensure at least to contain a specific quantity of items.
    /// </summary>
    /// <param name="count">Number of items to be at least in the collections.</param>
    /// <param name="value">Optional. String to be padded if number of items is less than the count.</param>
    /// <remarks>Padding will be performed at the start of the collection. If there are more items in the collection than the count, the collection is left unchanged.</remarks>
    public virtual void PadStart(int count, string value = "")
    {
      bool bSkip = false;

      if (String.IsNullOrEmpty(value) && this.IgnoreEmpty)
      {
        bSkip = true;
      }

      if (!bSkip)
      {
        while (base.List.Count < count)
        {
          base.List.Insert(0, value);
          RaiseItemsChangedEvent(value, StringCollectionOperation.Insert);
        }
      }
    }

    /// <summary>
    /// Pads the string collection to ensure at least to contain a specific quantity of items.
    /// </summary>
    /// <param name="count">Number of items to be at least in the collections.</param>
    /// <param name="value">Optional. String to be padded if number of items is less than the count.</param>
    /// <remarks>Padding will be performed at the end of the collection. If there are more items in the collection than the count, the collection is left unchanged.</remarks>
    public virtual void PadEnd(int count, string value = "")
    {
      bool bSkip = false;

      if (String.IsNullOrEmpty(value) && this.IgnoreEmpty)
      {
        bSkip = true;
      }

      if (!bSkip)
      {
        while (base.List.Count < count)
        {
          Add(value);
          RaiseItemsChangedEvent(value, StringCollectionOperation.Insert);
        }
      }
    }

    /// <summary>
    /// Retrieves the index of the referred string stored the string collection.
    /// </summary>
    /// <param name="value">Reference of a stored string in the string  collection.</param>
    /// <returns>Ordinal number of matched string or -1 if not exists in the collection.</returns>
    public int IndexOf(string value)
    {
      return base.List.IndexOf(value);
    }

    /// <summary>
    /// Gets the first string stored in the string collection.
    /// </summary>
    /// <returns>First string in the string collection.</returns>
    /// <remarks>Will raise an exception if the collection is empty.</remarks>
    public string First()
    {
      return (string)base.List[0];
    }

    /// <summary>
    /// Gets the first string stored in the string collection, or its default value if its empty.
    /// </summary>
    /// <returns>First string in the string collection or <c>null</c> if collection is empty.</returns>
    public string? FirstOrDefault()
    {
      return base.List.Count == 0 ? null : (string)base.List[0];
    }

    /// <summary>
    /// Gets the last string stored in the string collection.
    /// </summary>
    /// <returns>Last string in the string collection.</returns>
    /// <remarks>Will raise an exception if the collection is empty.</remarks>
    public string Last()
    {
      return (string)base.List[base.List.Count - 1];
    }

    /// <summary>
    /// Gets the last string stored in the string collection, or its default value if its empty
    /// </summary>
    /// <returns>Last string in the string collection or <c>null</c> if collection is empty</returns>
    public string? LastOrDefault()
    {
      return base.List.Count == 0 ? null : (string)base.List[base.List.Count - 1];
    }

    /// <summary>
    /// Determines whether the specified string (case sensitive) is in the StringCollection.
    /// </summary>
    /// <param name="value">The string to locate in the StringCollection. The value can be a <c>null</c>-reference</param>
    /// <returns><c>True</c> if value is found in the StringCollection, otherwise <c>False</c></returns>
    public bool Contains(string value)
    {
      return (base.List.IndexOf(value) != -1);
    }

    /// <summary>
    /// Determines whether the specified string (by comparison type) is in the StringCollection 
    /// </summary>
    /// <param name="value">The string to locate in the StringCollection. The value can be a <c>null</c>-reference</param>
    /// <param name="comparisonType">String comparison type to use</param>
    /// <returns><c>True</c> if value is found in the StringCollection, otherwise <c>False</c></returns>
    public bool Contains(string value, StringComparison comparisonType)
    {
      bool Result = false;

      foreach (string sItem in base.List)
      {
        if (String.Compare(sItem, value, comparisonType) == 0)
        {
          Result = true;
          break;
        }
      }

      return Result;
    }

    /// <summary>
    /// Checks if the content of a specified collection is equal to the current content of this collection
    /// </summary>
    /// <param name="collection">String collection to check.</param>
    /// <returns><c>True</c> if the content of the specified collection is equal to the content of the current collection,  otherwise <c>False</c></returns>
    /// <remarks>The check of the content is case-sensetive.</remarks>
    public bool IsContentEqual(StringCollection collection)
    {
      // check if collection is set
      if (collection == null)
      {
        return false;
      }

      // check number of elements
      if (base.List.Count != collection.Count)
      {
        return false;
      }

      // check content
      foreach (string sItem in base.List)
      {
        if (!collection.Contains(sItem, StringComparison.Ordinal))
        {
          return false;
        }
      }

      return true;
    }

    /// <summary>
    /// Gets the intersection with a specified string collection.
    /// </summary>
    /// <param name="collection">String collection to intersection for.</param>
    /// <returns>Intersection with a specified string collection.</returns>
    public StringCollection Intersection(StringCollection collection)
    {
      StringCollection Result = new StringCollection();

      if (collection == null)
      {
        Result.Assign(this);
      }
      else
      {
        // get all intersecting elements
        StringCollection colSource;
        StringCollection colDestination;

        if (this.Count > collection.Count)
        {
          colSource = collection;
          colDestination = this;
        }
        else
        {
          colSource = this;
          colDestination = collection;
        }

        foreach (string sItem in colSource)
        {
          if (colDestination.Contains(sItem, StringComparison.Ordinal))
          {
            Result.AddDistinct(sItem);
          }
        }
      }

      return Result;
    }

    /// <summary>
    /// Gets the intersection with a specified string collection.
    /// </summary>
    /// <param name="collection">String collection to intersection for.</param>
    /// <param name="comparisonType">Comparison type for checking collection items.</param>
    /// <returns>Intersection with a specified string collection.</returns>
    public StringCollection Intersection(StringCollection collection, StringComparison comparisonType)
    {
      StringCollection Result = new StringCollection();

      if (collection == null)
      {
        Result.Assign(this);
      }
      else
      {
        // get all intersecting elements
        StringCollection colSource;
        StringCollection colDestination;

        if (this.Count > collection.Count)
        {
          colSource = collection;
          colDestination = this;
        }
        else
        {
          colSource = this;
          colDestination = collection;
        }

        foreach (string sItem in colSource)
        {
          if (colDestination.Contains(sItem, comparisonType))
          {
            Result.AddDistinct(sItem);
          }
        }
      }

      return Result;
    }

    /// <summary>
    /// Adds a new string to the string collection at a specific position.
    /// </summary>
    /// <param name="index">Ordinal number of a string where the new entry is inserted into the string collection <strong>in front of</strong> the current one.</param>
    /// <param name="value">Reference to the string item.</param>
    public void Insert(int index, string? value)
    {
      if (String.IsNullOrEmpty(value))
      {
        if (!(value == null || this.IgnoreEmpty))
        {
          base.List.Insert(index, value);

          RaiseItemsChangedEvent(value, StringCollectionOperation.Insert);
        }
        else
        {
          base.List.Insert(index, value);

          RaiseItemsChangedEvent(value, StringCollectionOperation.Insert);
        }
      }
    }

    /// <summary>
    /// Adds a new formatted string to the string collection at a specific position.
    /// </summary>
    /// <param name="index">Ordinal number of a string where the new entry is inserted into the string collection <strong>in front of</strong> the current one.</param>
    /// <param name="format">A composite format string.</param>
    /// <param name="args">An System.Object array containing zero or more objects to format.</param>
    public void InsertFormat(int index, string format, params object[] args)
    {
      Insert(index, String.Format(CultureInfo.InvariantCulture, format, args));
    }

    /// <summary>
    /// Copies the elements of a StringCollection to the beginning of this
    /// StringCollection. 
    /// </summary>
    /// <param name="values">String collection to insert.</param>
    /// <param name="index">Ordinal number of a string where the new entry collection is inserted into the string collection <strong>in front of</strong> the current one.</param>
    public virtual void InsertRange(int index, StringCollection values)
    {
      foreach (string sItem in values)
      {
        Insert(index++, sItem);
      }
    }

    /// <summary>
    /// Removes a string from the string collection.
    /// </summary>
    /// <param name="value">Reference of a string stored in the node list collection.</param>
    /// <param name="ignoreNonExistingValue">Optional. Indicator whether non-existing value will be ignored or not. Default value is <strong>true</strong>.</param>
    public virtual void Remove(string value, bool ignoreNonExistingValue = true)
    {
      if (ignoreNonExistingValue)
      {
        if (this.Contains(value, StringComparison.Ordinal))
        {
          RaiseItemsChangedEvent(value, StringCollectionOperation.Remove);

          base.List.Remove(value);
        }
        else
        {
          RaiseItemsChangedEvent(value, StringCollectionOperation.Remove);

          base.List.Remove(value);
        }
      }
    }

    /// <summary>
    /// Removes a string from the string collection.
    /// </summary>
    /// <param name="value">Reference of a string stored in the node list collection.</param>
    /// <param name="comparisonType">Comparison type for value look-up.</param>
    /// <param name="ignoreNonExistingValue">Optional. Indicator whether non-existing value will be ignored or not. Default value is <strong>true</strong>.</param>
    public virtual void Remove(string value, StringComparison comparisonType, bool ignoreNonExistingValue = true)
    {
      if (ignoreNonExistingValue)
      {
        if (this.Contains(value, comparisonType))
        {
          RaiseItemsChangedEvent(value, StringCollectionOperation.Remove);

          base.List.Remove(value);
        }
        else
        {
          RaiseItemsChangedEvent(value, StringCollectionOperation.Remove);

          base.List.Remove(value);
        }
      }
    }

    /// <summary>
    /// Removes the elements of a string array from the StringCollection. 
    /// </summary>
    /// <param name="values">String array to be removed.</param>
    /// <param name="ignoreNonExistingValue">Optional. Indicator whether non-existing value will be ignored or not. Default value is <strong>true</strong>.</param>
    public virtual void RemoveRange(string[] values, bool ignoreNonExistingValue = true)
    {
      foreach (string sItem in values)
      {
        Remove(sItem, ignoreNonExistingValue);
      }
    }

    /// <summary>
    /// Removes the elements of a string array from the StringCollection. 
    /// </summary>
    /// <param name="values">String array to be removed.</param>
    /// <param name="comparisonType">Comparison type for value look-up.</param>
    /// <param name="ignoreNonExistingValue">Optional. Indicator whether non-existing value will be ignored or not. Default value is <strong>true</strong>.</param>
    public virtual void RemoveRange(string[] values, StringComparison comparisonType, bool ignoreNonExistingValue = true)
    {
      foreach (string sItem in values)
      {
        Remove(sItem, comparisonType, ignoreNonExistingValue);
      }
    }

    /// <summary>
    /// Removes the elements of a StringCollection from the StringCollection. 
    /// </summary>
    /// <param name="values">StringCollection to be removed.</param>
    /// <param name="ignoreNonExistingValue">Optional. Indicator whether non-existing value will be ignored or not. Default value is <strong>true</strong>.</param>
    public virtual void RemoveRange(StringCollection values, bool ignoreNonExistingValue = true)
    {
      foreach (string sItem in values)
      {
        Remove(sItem, ignoreNonExistingValue);
      }
    }

    /// <summary>
    /// Removes the elements of a StringCollection from the StringCollection. 
    /// </summary>
    /// <param name="values">StringCollection to be removed.</param>
    /// <param name="comparisonType">Comparison type for value look-up.</param>
    /// <param name="ignoreNonExistingValue">Optional. Indicator whether non-existing value will be ignored or not. Default value is <strong>true</strong>.</param>
    public virtual void RemoveRange(StringCollection values, StringComparison comparisonType, bool ignoreNonExistingValue = true)
    {
      foreach (string sItem in values)
      {
        Remove(sItem, comparisonType, ignoreNonExistingValue);
      }
    }

    /// <summary>
    /// Removes all null (nothing) or empty entries from the StringCollection.
    /// </summary>
    public void RemoveNullOrEmpty()
    {
      for (int i = (base.List.Count - 1); i >= 0; i--)
      {
        if (String.IsNullOrEmpty((string)base.List[i]))
        {
          RaiseItemsChangedEvent((string)base.List[i], StringCollectionOperation.Remove);

          base.List.RemoveAt(i);
        }
      }
    }

    /// <summary>
    /// Sort the string collection either ascending or descending.
    /// </summary>
    /// <param name="direction">Direction of sort.</param>
    internal void Sort(SortDirection direction)
    {
      this.SortComparer.SortDirection = direction;
      this.InnerList.Sort(this.SortComparer);
    }

    /// <summary>
    /// Returns a string that represents the collection.
    /// </summary>
    /// <param name="delimeter">Optional. Separator for each string item. Default value is &quot; &quot; (space)</param>
    /// <returns>Formatted string that represents the string collection.</returns>
    /// <remarks>If parameter <strong>Delimeter</strong> is set to Nothing or to an empty value the <strong>ToString</strong> function of the base class is called instead.</remarks>
    public string ToString(string delimeter = " ")
    {
      return ToString(0, this.Count, delimeter);
    }

    /// <summary>
    /// Returns a string that represents the collection.
    /// </summary>
    /// <param name="startIndex">Start index to begin with.</param>
    /// <param name="count">Number of items to work on.</param>
    /// <param name="delimeter">Optional. Separator for each string item. Default value is &quot; &quot; (space)</param>
    /// <returns>Formatted string that represents the string collection.</returns>
    /// <remarks>If parameter <strong>Delimeter</strong> is set to Nothing or to an empty value the <strong>ToString</strong> function of the base class is called instead.</remarks>
    public string ToString(int startIndex, int count, string delimeter = " ")
    {
      if (String.IsNullOrEmpty(delimeter))
      {
        return base.ToString();
      }
      else
      {
        StringBuilder sbResult = new StringBuilder();
        int iStartIndex = Math.Max(0, startIndex);
        int iCount = Math.Min(this.Count, count);

        for (int i = iStartIndex; i < iCount; i++)
        {
          if (i == iStartIndex)
          {
            sbResult.Append(this[i]);
          }
          else
          {
            sbResult.Append(delimeter + this[i]);
          }
        }

        return sbResult.ToString();
      }
    }

    /// <summary>
    /// Returns a string that represents the collection. Each item is enclosed by a specific quotation.
    /// </summary>
    /// <param name="delimeter">Separator for each string item.</param>
    /// <param name="quotation">Quotation that encloses each string item.</param>
    /// <returns>Formatted string that represents the string collection.</returns>
    public string ToString(string delimeter, string quotation)
    {
      return ToString(0, this.Count, delimeter, quotation);
    }

    /// <summary>
    /// Returns a string that represents the collection. Each item is enclosed by a specific quotation.
    /// </summary>
    /// <param name="startIndex">Start index to begin with.</param>
    /// <param name="count">Number of items to work on.</param>
    /// <param name="delimeter">Separator for each string item.</param>
    /// <param name="quotation">Quotation that encloses each string item.</param>
    /// <returns>Formatted string that represents the string collection.</returns>
    public string ToString(int startIndex, int count, string delimeter, string quotation)
    {
      StringBuilder sbResult = new StringBuilder();
      int iStartIndex = Math.Max(0, startIndex);
      int iCount = Math.Min(this.Count, count);

      for (int i = iStartIndex; i < iCount; i++)
      {
        if (i == iStartIndex)
        {
          sbResult.AppendFormat(CultureInfo.InvariantCulture, "{0}{1}{0}", quotation, this[i]);
        }
        else
        {
          sbResult.AppendFormat(CultureInfo.InvariantCulture, "{0}{1}{2}{1}", delimeter, quotation, this[i]);
        }
      }

      return sbResult.ToString();
    }

    /// <summary>
    /// Returns a string that represents the collection. Each item is enclosed by a specific quotation (beginning and ending).
    /// </summary>
    /// <param name="delimeter">Separator for each string item.</param>
    /// <param name="beginQuotation">Beginning quotation that encloses each string item.</param>
    /// <param name="endQuotation">Ending quotation that encloses each string item.</param>
    /// <returns>Formatted string that represents the string collection.</returns>
    public string ToString(string delimeter, string beginQuotation, string endQuotation)
    {
      return ToString(0, this.Count, delimeter, beginQuotation, endQuotation);
    }

    /// <summary>
    /// Returns a string that represents the collection. Each item is enclosed by a specific quotation (beginning and ending).
    /// </summary>
    /// <param name="startIndex">Start index to begin with.</param>
    /// <param name="count">Number of items to work on.</param>
    /// <param name="delimeter">Separator for each string item.</param>
    /// <param name="beginQuotation">Beginning quotation that encloses each string item.</param>
    /// <param name="endQuotation">Ending quotation that encloses each string item.</param>
    /// <returns>Formatted string that represents the string collection.</returns>
    public string ToString(int startIndex, int count, string delimeter, string beginQuotation, string endQuotation)
    {
      StringBuilder sbResult = new StringBuilder();
      int iStartIndex = Math.Max(0, startIndex);
      int iCount = Math.Min(this.Count, count);

      for (int i = iStartIndex; i < iCount; i++)
      {
        if (i == iStartIndex)
        {
          sbResult.AppendFormat(CultureInfo.InvariantCulture, "{0}{1}{2}", beginQuotation, this[i], endQuotation);
        }
        else
        {
          sbResult.AppendFormat(CultureInfo.InvariantCulture, "{0}{1}{2}{3}", delimeter, beginQuotation, this[i], endQuotation);
        }
      }

      return sbResult.ToString();
    }

    /// <summary>
    /// Gets a string array representing all items in this collection.
    /// </summary>
    /// <returns>String array representing all items in this collection.</returns>
    public string[] ToArray()
    {
      string[] Result = new string[this.Count];

      for (int i = 0; i < this.Count; i++)
      {
        Result[i] = this[i];
      }

      return Result;
    }

    /// <summary>
    /// Gets a string array representing specific items in this collection.
    /// </summary>
    /// <param name="startIndex">Start index to begin with.</param>
    /// <param name="count">Number of items to work on.</param>
    /// <returns>String array representing <c>count</c> items beginning at the startindex in this collection.</returns>
    public string[] ToArray(int startIndex, int count)
    {
      int iStartIndex = Math.Max(0, startIndex);
      int iCount = Math.Min(this.Count, count);

      string[] Result = new string[iCount];

      for (int i = iStartIndex; i < iCount; i++)
      {
        Result[i] = this[i];
      }

      return Result;
    }

    /// <summary>
    /// Gets a generic string list representing all items in this collection.
    /// </summary>
    /// <returns>Generic string list representing all items in this collection.</returns>
    public List<string> ToList()
    {
      return ToList(0, this.Count);
    }

    /// <summary>
    /// Gets a generic string list representing specific items in this collection.
    /// </summary>
    /// <param name="startIndex">Start index to begin with.</param>
    /// <param name="count">Number of items to work on.</param>
    /// <returns>Generic string list representing <c>count</c> items beginning at the start index in this collection.</returns>
    public List<string> ToList(int startIndex, int count)
    {
      List<string> Result = new List<string>();
      int iStartIndex = Math.Max(0, startIndex);
      int iCount = Math.Min(this.Count, count);

      for (int i = iStartIndex; i < iCount; i++)
      {
        Result.Add(this[i]);
      }

      return Result;
    }

    /// <summary>
    /// Returns a new collection containing a specified collection of strings. 
    /// </summary>
    /// <param name="expression">String expression containing substrings and delimiters.</param>
    /// <param name="delimeter">Any single character used to identify substring limits. If Delimiter is omitted, the space character (&quot; &quot;) is assumed to be the delimiter.</param>
    /// <param name="startIndex">Start index to begin the separation. If omitted first substring is assumed.</param>
    /// <param name="count">Maximum number of substrings into which the input string should be split. The default, –1, indicates that the input string should be split at every occurrence of the delimiter string.
    /// </param>
    /// <returns>List of strings. If Expression is a zero-length string (&quot;&quot;), Split returns a single-element list containing a zero-length string. If Delimiter is a zero-length string, or if it
    /// does not appear anywhere in Expression, Split returns a single-element list containing the entire Expression string.</returns>
    /// <remarks>Each substring will be trimmed.</remarks>
    public static StringCollection Split(string expression, char delimeter = ' ', int startIndex = 1, int count = -1)
    {
      // check if start index is valid
      if (startIndex <= 0)
      {
        throw new ArgumentOutOfRangeException(nameof(startIndex), startIndex, String.Format(CultureInfo.InvariantCulture, Resources.StartIndexOutOfRangeMustBeValueOrGreaterFormatErrorMessage, startIndex));
      }

      // check if count is valid
      if (count <= -2)
      {
        throw new ArgumentOutOfRangeException(nameof(count), count, String.Format(CultureInfo.InvariantCulture, Resources.CountOutOfRangeMustBeValueOrGreaterFormatErrorMessage, count));
      }

      StringCollection Result = new StringCollection();

      if (!String.IsNullOrEmpty(expression))
      {
        char[] aChars = expression.ToCharArray();
        StringBuilder sbEntry = new StringBuilder();
        int iIndex = 1;
        int iCount = 0;

        foreach (char cChar in aChars)
        {
          if (cChar == delimeter)
          {
            if (iIndex >= startIndex)
            {
              if (count == -1 || iCount < count)
              {
                Result.Add(sbEntry.ToString().Trim());

                sbEntry = new StringBuilder();

                iIndex++;
                iCount++;
              }
            }
          }
          else
          {
            sbEntry.Append(cChar);
          }
        }

        // add remaining part to the result
        Result.Add(sbEntry.ToString().Trim());
      }

      return Result;
    }

    /// <summary>
    /// Returns a new collection containing a specified collection of distinct strings. 
    /// </summary>
    /// <param name="expression">String expression containing substrings and delimiters.</param>
    /// <param name="delimeter">Any single character used to identify substring limits. If Delimiter is omitted, the space character (&quot; &quot;) is assumed to be the delimiter.</param>
    /// <param name="startIndex">Start index to begin the separation. If omitted first substring is assumed.</param>
    /// <param name="count">Maximum number of substrings into which the input string should be split. The default, –1, indicates that the input string should be split at every occurrence of the delimiter string.
    /// </param>
    /// <returns>List of strings. If Expression is a zero-length string (&quot;&quot;), Split returns a single-element list containing a zero-length string. If Delimiter is a zero-length string, or if it
    /// does not appear anywhere in Expression, Split returns a single-element list containing the entire Expression string.</returns>
    /// <remarks>Each substring will be trimmed. Duplicate substrings will be ignored.</remarks>
    public static StringCollection SplitDistinct(string expression, char delimeter = ' ', int startIndex = 1, int count = -1)
    {
      // check if start index is valid
      if (startIndex <= 0)
      {
        throw new ArgumentOutOfRangeException(nameof(startIndex), startIndex, String.Format(CultureInfo.InvariantCulture, Resources.StartIndexOutOfRangeMustBeValueOrGreaterFormatErrorMessage, startIndex));
      }

      // check if count is valid
      if (count <= -2)
      {
        throw new ArgumentOutOfRangeException(nameof(count), count, String.Format(CultureInfo.InvariantCulture, Resources.CountOutOfRangeMustBeValueOrGreaterFormatErrorMessage, count));
      }

      StringCollection Result = new StringCollection();

      if (!String.IsNullOrEmpty(expression))
      {
        char[] aChars = expression.ToCharArray();
        StringBuilder sbEntry = new StringBuilder();
        int iIndex = 1;
        int iCount = 0;

        foreach (char cChar in aChars)
        {
          if (cChar == delimeter)
          {
            if (iIndex >= startIndex)
            {
              if (count == -1 || iCount < count)
              {
                Result.Add(sbEntry.ToString().Trim());
                sbEntry = new StringBuilder();

                iIndex++;
                iCount++;
              }
            }
          }
          else
          {
            sbEntry.Append(cChar);
          }
        }

        // add remaining distinct part to the result
        Result.AddDistinct(sbEntry.ToString().Trim());
      }

      return Result;
    }

    /// <summary>
    /// Converts string collection to a single string. The separator is the delimiter for each item with the collection. The collection is left untouched.
    /// </summary>
    /// <param name="delimeter">Optional. The delimiter for the output string. Default value is a space character.</param>
    /// <param name="startIndex">Optional. Start index of the collection.</param>
    /// <param name="count">Optional. Limits the number of collection-items to process. -1 uses entire collection. Default value is <c>-1</c>.</param>
    /// <returns>String that contains each collection item separated with a specific delimiter.</returns>
    public string Join(char delimeter = ' ', int startIndex = 0, int count = -1)
    {
      // check if collection is empty
      if (this.Count == 0)
      {
        return "";
      }

      // check if start index is valid
      if (startIndex < 0 || startIndex > this.Count)
      {
        throw new ArgumentOutOfRangeException(nameof(startIndex), startIndex, String.Format(CultureInfo.InvariantCulture, Resources.StartIndexOutOfRangeMustBeBetweenValuesFormatErrorMessage, 0, Math.Max(this.Count - 1, 0)));
      }

      // check if count is valid
      if (count != -1 && (count < 0 || startIndex + count >= this.Count))
      {
        throw new ArgumentOutOfRangeException(nameof(count), count, String.Format(CultureInfo.InvariantCulture, Resources.CountOutOfRangeMustBeValueOrGreaterFormatErrorMessage, -1));
      }

      StringBuilder Result = new StringBuilder();
      int iUpperBound;

      if (count == -1)
      {
        iUpperBound = this.Count - 1;
      }
      else
      {
        iUpperBound = startIndex + (count - 1);
      }

      for (int i = startIndex; i <= iUpperBound; i++)
      {
        if (Result.Length == 0)
        {
          Result.Append(this[i]);
        }
        else
        {
          Result.Append(delimeter + this[i]);
        }
      }

      return Result.ToString();
    }

    /// <summary>
    /// Raise event upon changed item
    /// </summary>
    /// <param name="value">Value of changed item</param>
    /// <param name="operation">Operation of change</param>
    private void RaiseItemsChangedEvent(string? value, StringCollectionOperation operation)
    {
      this.ItemsChanged?.Invoke(this, new StringCollectionEventArgs(value, operation));
    }

    /// <summary>
    /// Indicator if object was already disposed by class itself.
    /// </summary>
    private bool _Disposed;
  }
}
