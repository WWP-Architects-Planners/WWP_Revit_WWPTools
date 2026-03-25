using System;
using Autodesk.Revit.UI;
class P {
  static void Main() {
    foreach (var name in Enum.GetNames(typeof(PostableCommand))) {
      if (name.IndexOf("Group", StringComparison.OrdinalIgnoreCase) >= 0 ||
          name.IndexOf("Sketch", StringComparison.OrdinalIgnoreCase) >= 0 ||
          name.IndexOf("Boundary", StringComparison.OrdinalIgnoreCase) >= 0 ||
          name.IndexOf("Profile", StringComparison.OrdinalIgnoreCase) >= 0 ||
          name.IndexOf("Selection", StringComparison.OrdinalIgnoreCase) >= 0) {
        Console.WriteLine(name);
      }
    }
  }
}
