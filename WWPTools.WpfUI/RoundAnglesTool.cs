using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using WWPTools.WpfUI.Views;

namespace WWPTools.WpfUI
{
    public static class RoundAnglesLauncher
    {
        private static RoundAnglesController _controller;

        public static void Show(object uiAppObject)
        {
            var uiApp = uiAppObject as UIApplication;
            if (uiApp == null)
                throw new ArgumentException("Expected Autodesk.Revit.UI.UIApplication", nameof(uiAppObject));

            if (_controller != null && _controller.IsAlive)
            {
                _controller.ShowWindow();
                _controller.RequestRefresh();
                return;
            }

            _controller = new RoundAnglesController(uiApp);
            _controller.ShowWindow();
            _controller.RequestRefresh();
        }
    }

    internal sealed class RoundAnglesController
    {
        private const string Title = "Round Off-Axis Sketch Lines";
        private const string OffAxisWarningCoreText = "slightly off axis";
        private const string OffAxisWarningAccuracyText = "may cause inaccuracies";
        private const int AnglePrecision = 2;
        private const double AngleEpsilonDegrees = 1e-6;

        private readonly UIApplication _uiApp;
        private readonly RoundAnglesExternalEventHandler _handler;
        private readonly ExternalEvent _externalEvent;
        private readonly RoundAnglesWindow _window;
        private readonly object _requestLock = new object();
        private RoundAnglesRequest _pendingRequest;
        private string _lastLogPath = "";
        private int? _activeGroupContextId;
        private bool _isClosed;

        public RoundAnglesController(UIApplication uiApp)
        {
            _uiApp = uiApp;
            _handler = new RoundAnglesExternalEventHandler(this);
            _externalEvent = ExternalEvent.Create(_handler);
            _window = new RoundAnglesWindow
            {
                Controller = this
            };
            _window.SetOwnerHandle(uiApp.MainWindowHandle);
        }

        public bool IsAlive => !_isClosed && _window.IsLoaded;

        public void ShowWindow()
        {
            if (_window.IsVisible)
            {
                _window.Activate();
                return;
            }

            _window.Show();
        }

        public void NotifyWindowClosed()
        {
            _isClosed = true;
        }

        public void RequestRefresh()
        {
            QueueRequest(new RoundAnglesRequest(RoundAnglesRequestKind.Refresh));
        }

        public void RequestApply(IList<RoundAnglesItem> items)
        {
            QueueRequest(new RoundAnglesRequest(RoundAnglesRequestKind.Apply) { Items = items?.ToList() ?? new List<RoundAnglesItem>() });
        }

        public void RequestSelect(RoundAnglesItem item)
        {
            QueueRequest(new RoundAnglesRequest(RoundAnglesRequestKind.Select) { Item = item });
        }

        public void RequestGoToView(RoundAnglesItem item)
        {
            QueueRequest(new RoundAnglesRequest(RoundAnglesRequestKind.GoToView) { Item = item });
        }

        public void RequestEditGroup(RoundAnglesItem item)
        {
            QueueRequest(new RoundAnglesRequest(RoundAnglesRequestKind.EditGroup) { Item = item });
        }

        public void RequestEditBoundary(RoundAnglesItem item)
        {
            QueueRequest(new RoundAnglesRequest(RoundAnglesRequestKind.EditBoundary) { Item = item });
        }

        public void RequestZoomTo(RoundAnglesItem item)
        {
            QueueRequest(new RoundAnglesRequest(RoundAnglesRequestKind.ZoomTo) { Item = item });
        }

        public void RequestIsolate(RoundAnglesItem item)
        {
            QueueRequest(new RoundAnglesRequest(RoundAnglesRequestKind.Isolate) { Item = item });
        }

        internal void Execute(UIApplication uiApp)
        {
            var request = TakeRequest();
            if (request == null)
                return;

            switch (request.Kind)
            {
                case RoundAnglesRequestKind.Refresh:
                    ExecuteRefresh(uiApp);
                    break;
                case RoundAnglesRequestKind.Apply:
                    ExecuteApply(uiApp, request.Items);
                    break;
                case RoundAnglesRequestKind.Select:
                    ExecuteSelect(uiApp, request.Item);
                    break;
                case RoundAnglesRequestKind.GoToView:
                    ExecuteGoToView(uiApp, request.Item);
                    break;
                case RoundAnglesRequestKind.EditGroup:
                    ExecuteEditGroup(uiApp, request.Item);
                    break;
                case RoundAnglesRequestKind.EditBoundary:
                    ExecuteEditBoundary(uiApp, request.Item);
                    break;
                case RoundAnglesRequestKind.ZoomTo:
                    ExecuteZoomTo(uiApp, request.Item);
                    break;
                case RoundAnglesRequestKind.Isolate:
                    ExecuteIsolate(uiApp, request.Item);
                    break;

            }
        }

        private void QueueRequest(RoundAnglesRequest request)
        {
            if (_isClosed)
                return;

            lock (_requestLock)
            {
                _pendingRequest = request;
            }

            _externalEvent.Raise();
        }

        private RoundAnglesRequest TakeRequest()
        {
            lock (_requestLock)
            {
                var request = _pendingRequest;
                _pendingRequest = null;
                return request;
            }
        }

        private void ExecuteRefresh(UIApplication uiApp)
        {
            var doc = uiApp.ActiveUIDocument?.Document;
            if (doc == null)
            {
                _window.ShowError("No active Revit document found.", Title);
                return;
            }

            var editableGroupId = ResolveEditableGroupId(uiApp.ActiveUIDocument);
            var warnings = CollectTargetWarnings(doc);
            var skipped = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
            var items = new List<RoundAnglesItem>();

            int index = 0;
            foreach (var warning in warnings)
            {
                index += 1;
                BuildItemForWarning(doc, warning, index, items, skipped, editableGroupId);
            }

            items.Sort((left, right) =>
            {
                var deltaCompare = Math.Abs(right.DeltaDegrees).CompareTo(Math.Abs(left.DeltaDegrees));
                if (deltaCompare != 0)
                    return deltaCompare;
                return string.Compare(left.DisplayText, right.DisplayText, StringComparison.OrdinalIgnoreCase);
            });

            _window.LoadItems(warnings.Count, FormatSkipped(skipped), items, _lastLogPath);
        }

        private void ExecuteApply(UIApplication uiApp, IList<RoundAnglesItem> items)
        {
            var doc = uiApp.ActiveUIDocument?.Document;
            if (doc == null)
            {
                _window.ShowError("No active Revit document found.", Title);
                return;
            }

            var applicable = (items ?? Array.Empty<RoundAnglesItem>())
                .Where(item => item != null && item.CanApply)
                .ToList();

            if (applicable.Count == 0)
            {
                var totalRows = (items ?? Array.Empty<RoundAnglesItem>()).Count(item => item != null);
                var groupedRows = (items ?? Array.Empty<RoundAnglesItem>()).Count(item => item != null && item.Status == RoundAnglesItemStatus.GroupedSkip);
                var failedRows = (items ?? Array.Empty<RoundAnglesItem>()).Count(item => item != null && item.Status == RoundAnglesItemStatus.Failed);
                var message = totalRows == 0
                    ? "There are no rows in the current list. Refresh to rescan the model."
                    : string.Format(
                        CultureInfo.InvariantCulture,
                        "There are no pending rows left in the current list. Remaining visible rows: {0}. Failed: {1}. Grouped skips: {2}. Refresh to rescan the model warnings.",
                        totalRows,
                        failedRows,
                        groupedRows);
                _window.ShowInfo(message, Title);
                return;
            }

            var updated = 0;
            var failed = 0;
            using (var transaction = new Transaction(doc, Title))
            {
                transaction.Start();
                foreach (var item in applicable)
                {
                    var element = GetElement(doc, item.ElementIdValue);
                    if (element == null)
                    {
                        item.MarkFailed("Element no longer exists.");
                        failed += 1;
                        continue;
                    }

                    var repin = false;
                    try
                    {
                        if (item.WasPinned)
                        {
                            try
                            {
                                element.Pinned = false;
                                repin = true;
                            }
                            catch
                            {
                                repin = false;
                            }
                        }

                        var locationCurve = element.Location as LocationCurve;
                        if (locationCurve == null)
                            throw new InvalidOperationException("Location curve unavailable during apply.");

                        locationCurve.Curve = item.RotatedCurve;
                        item.MarkUpdated();
                        updated += 1;
                    }
                    catch (Exception ex)
                    {
                        item.MarkFailed(ex.Message);
                        failed += 1;
                    }
                    finally
                    {
                        if (repin)
                        {
                            try
                            {
                                element.Pinned = true;
                            }
                            catch
                            {
                            }
                        }
                    }
                }

                transaction.Commit();
            }

            _lastLogPath = WriteLog(doc, applicable, updated, failed);
            _window.RefreshPresentation(_lastLogPath);
        }

        private void ExecuteSelect(UIApplication uiApp, RoundAnglesItem item)
        {
            SelectElement(uiApp.ActiveUIDocument, item?.PreferredUiElementIdValue);
        }

        private void ExecuteGoToView(UIApplication uiApp, RoundAnglesItem item)
        {
            var doc = uiApp.ActiveUIDocument?.Document;
            if (doc == null || item == null)
                return;

            var targetElement = GetElement(doc, item.PreferredViewElementIdValue ?? item.PreferredUiElementIdValue);
            var view = GetOwnerView(doc, targetElement) ?? GetSameLevelView(doc, targetElement);
            if (view == null)
            {
                _window.ShowInfo("No good view could be found.", Title);
                return;
            }

            uiApp.ActiveUIDocument.ActiveView = view;
            SelectElement(uiApp.ActiveUIDocument, item.PreferredUiElementIdValue);
        }

        private void ExecuteEditGroup(UIApplication uiApp, RoundAnglesItem item)
        {
            var groupId = item?.GroupElementIdValue;
            if (!groupId.HasValue)
            {
                _window.ShowInfo("This row does not resolve to a group instance.", Title);
                return;
            }

            _activeGroupContextId = groupId;
            SelectElement(uiApp.ActiveUIDocument, groupId);

            // Enter edit mode if not already in it; if already inside, this is a no-op.
            TryPostNamedCommand(uiApp, "EditGroup");

            // Always refresh so items are re-evaluated with the updated group context.
            // This makes the button work as "Set Active Group + Refresh" even when
            // the user manually entered edit mode through the Revit UI.
            ExecuteRefresh(uiApp);
        }

        private void ExecuteEditBoundary(UIApplication uiApp, RoundAnglesItem item)
        {
            if (item == null)
                return;

            if (item.PreferredUiElementIdValue.HasValue)
                SelectElement(uiApp.ActiveUIDocument, item.PreferredUiElementIdValue);

            if (!TryPostNamedCommand(uiApp, "EditBoundary", "EditProfile", "EditSketch"))
                _window.ShowInfo("No boundary/profile edit command was available for this row.", Title);
        }

        private void ExecuteZoomTo(UIApplication uiApp, RoundAnglesItem item)
        {
            if (item == null)
                return;

            var uiDoc = uiApp.ActiveUIDocument;
            if (uiDoc == null)
                return;

            var preferredId = item.PreferredUiElementIdValue;
            if (preferredId.HasValue)
                SelectElement(uiDoc, preferredId);

            if (TryPostNamedCommand(uiApp, "SelectionBox"))
                return;

            try
            {
                if (preferredId.HasValue)
                {
                    uiDoc.ShowElements(new ElementId(preferredId.Value));
                    return;
                }
            }
            catch
            {
            }

            _window.ShowInfo("No good view could be found.", Title);
        }

        private void ExecuteIsolate(UIApplication uiApp, RoundAnglesItem item)
        {
            var uiDoc = uiApp.ActiveUIDocument;
            if (uiDoc == null || item?.PreferredUiElementIdValue == null)
                return;

            try
            {
                var ids = new Collection<ElementId> { new ElementId(item.PreferredUiElementIdValue.Value) };
                uiDoc.ActiveView.IsolateElementsTemporary(ids);
            }
            catch
            {
                _window.ShowInfo("Could not isolate this item in the active view.", Title);
            }
        }

        private static IList<object> CollectTargetWarnings(Document doc)
        {
            return (doc.GetWarnings()?.Cast<object>() ?? Enumerable.Empty<object>())
                .Where(warning => IsSupportedOffAxisWarning(WarningDescription(warning)))
                .ToList();
        }

        private static bool IsSupportedOffAxisWarning(string description)
        {
            var normalized = NormalizeText(description);
            if (string.IsNullOrWhiteSpace(normalized))
                return false;

            return normalized.IndexOf(OffAxisWarningCoreText, StringComparison.OrdinalIgnoreCase) >= 0
                && normalized.IndexOf(OffAxisWarningAccuracyText, StringComparison.OrdinalIgnoreCase) >= 0;
        }

        private void BuildItemForWarning(Document doc, object warning, int warningIndex, IList<RoundAnglesItem> items, IDictionary<string, int> skipped, int? editableGroupId)
        {
            var failingElements = WarningFailingElements(doc, warning);
            if (failingElements.Count == 0)
            {
                IncrementSkipped(skipped, "No failing elements");
                return;
            }

            var element = ResolveTargetElement(failingElements);
            if (element == null)
            {
                IncrementSkipped(skipped, "No supported line target");
                return;
            }

            if (IsPropertyLine(element))
            {
                IncrementSkipped(skipped, "Property line");
                return;
            }

            var contextElement = WarningContextElement(failingElements, element);
            var topParent = TopParentElement(doc, element, contextElement, null);
            var groupElement = GetGroupElement(doc, topParent) ?? GetGroupElement(doc, contextElement) ?? GetGroupElement(doc, element);

            if (IsGroupMember(element) && !editableGroupId.HasValue)
            {
                IncrementSkipped(skipped, "Group member");
                items.Add(RoundAnglesItem.CreateGroupedSkip(
                    warningIndex,
                    element,
                    contextElement,
                    groupElement,
                    WarningContextText(doc, failingElements, element)));
                return;
            }

            var locationCurve = element.Location as LocationCurve;
            if (locationCurve == null)
            {
                IncrementSkipped(skipped, "No location curve");
                return;
            }

            var line = locationCurve.Curve as Line;
            if (line == null)
            {
                IncrementSkipped(skipped, "Non-linear or vertical curve");
                return;
            }

            var currentAngle = CurveAngleDegrees(line);
            if (!currentAngle.HasValue)
            {
                IncrementSkipped(skipped, "Non-linear or vertical curve");
                return;
            }

            var targetAngle = NormalizeAngleDegrees(Math.Round(currentAngle.Value, AnglePrecision));
            var deltaDegrees = ShortestDeltaDegrees(currentAngle.Value, targetAngle);
            if (Math.Abs(deltaDegrees) <= AngleEpsilonDegrees)
            {
                IncrementSkipped(skipped, "Already rounded");
                return;
            }

            var rotatedCurve = RotateCurveAboutStart(line, DegreesToRadians(deltaDegrees));
            items.Add(RoundAnglesItem.CreatePending(
                warningIndex,
                element,
                contextElement,
                groupElement,
                WarningContextText(doc, failingElements, element),
                currentAngle.Value,
                targetAngle,
                deltaDegrees,
                line,
                rotatedCurve,
                element.Pinned));
        }

        private static Element ResolveTargetElement(IEnumerable<Element> failingElements)
        {
            var elements = (failingElements ?? Enumerable.Empty<Element>())
                .Where(element => element != null)
                .ToList();

            if (elements.Count == 0)
                return null;

            foreach (var element in elements.Where(IsPreferredOffAxisTarget))
                return element;

            foreach (var element in elements.Where(IsUsableOffAxisTarget))
                return element;

            return elements.LastOrDefault();
        }

        private static bool IsPreferredOffAxisTarget(Element element)
        {
            if (!IsUsableOffAxisTarget(element))
                return false;

            var categoryName = CategoryName(element);
            return string.Equals(categoryName, "Grids", StringComparison.OrdinalIgnoreCase)
                || string.Equals(categoryName, "Lines", StringComparison.OrdinalIgnoreCase)
                || string.Equals(categoryName, "Area Boundary", StringComparison.OrdinalIgnoreCase)
                || string.Equals(categoryName, "Area Boundary Lines", StringComparison.OrdinalIgnoreCase)
                || string.Equals(categoryName, "Sketch Lines", StringComparison.OrdinalIgnoreCase)
                || string.Equals(categoryName, "Model Lines", StringComparison.OrdinalIgnoreCase)
                || string.Equals(categoryName, "<Area Boundary>", StringComparison.OrdinalIgnoreCase);
        }

        private static bool IsUsableOffAxisTarget(Element element)
        {
            if (element == null || IsPropertyLine(element))
                return false;

            var locationCurve = element.Location as LocationCurve;
            var line = locationCurve?.Curve as Line;
            return CurveAngleDegrees(line).HasValue;
        }

        private static void IncrementSkipped(IDictionary<string, int> skipped, string reason)
        {
            if (string.IsNullOrWhiteSpace(reason))
                reason = "Unknown";

            skipped[reason] = skipped.TryGetValue(reason, out var count) ? count + 1 : 1;
        }

        private static string FormatSkipped(IDictionary<string, int> skipped)
        {
            if (skipped == null || skipped.Count == 0)
                return "Skipped: None";

            return "Skipped: " + string.Join(" | ", skipped.OrderBy(pair => pair.Key).Select(pair => string.Format("{0}: {1}", pair.Key, pair.Value)));
        }

        private static string WarningDescription(object warning)
        {
            try
            {
                if (warning == null)
                    return "";

                var method = warning.GetType().GetMethod("GetDescriptionText");
                if (method != null)
                    return method.Invoke(warning, null) as string ?? "";

                var property = warning.GetType().GetProperty("DescriptionText");
                return property?.GetValue(warning, null) as string ?? "";
            }
            catch
            {
                return "";
            }
        }

        private static IList<Element> WarningFailingElements(Document doc, object warning)
        {
            var elements = new List<Element>();
            try
            {
                if (warning == null)
                    return elements;

                var method = warning.GetType().GetMethod("GetFailingElements");
                var failingIds = method != null ? method.Invoke(warning, null) as System.Collections.IEnumerable : null;
                if (failingIds == null)
                    return elements;

                foreach (var elementIdObject in failingIds)
                {
                    var elementId = elementIdObject as ElementId;
                    if (elementId == null)
                        continue;

                    var element = doc.GetElement(elementId);
                    if (element != null)
                        elements.Add(element);
                }
            }
            catch
            {
            }

            return elements;
        }

        private static Element WarningContextElement(IEnumerable<Element> failingElements, Element targetElement)
        {
            var targetId = ElementIdValue(targetElement?.Id);
            foreach (var element in failingElements ?? Enumerable.Empty<Element>())
            {
                if (ElementIdValue(element?.Id) == targetId)
                    continue;
                return element;
            }

            return null;
        }

        private static string WarningContextText(Document doc, IEnumerable<Element> failingElements, Element targetElement)
        {
            var targetId = ElementIdValue(targetElement?.Id);
            var labels = new List<string>();

            foreach (var element in failingElements ?? Enumerable.Empty<Element>())
            {
                if (element == null || ElementIdValue(element.Id) == targetId)
                    continue;

                var label = ContextLabel(doc, element);
                if (!string.IsNullOrWhiteSpace(label) && !labels.Contains(label))
                    labels.Add(label);
            }

            if (labels.Count == 0)
            {
                var extras = new List<string>();
                var level = LevelName(doc, targetElement);
                var view = OwnerViewName(doc, targetElement);
                if (!string.IsNullOrWhiteSpace(level))
                    extras.Add("Level: " + level);
                if (!string.IsNullOrWhiteSpace(view))
                    extras.Add("View: " + view);
                return string.Join(" | ", extras);
            }

            if (labels.Count <= 3)
                return string.Join(" | ", labels);

            return string.Join(" | ", labels.Take(3)) + " | +" + (labels.Count - 3).ToString(CultureInfo.InvariantCulture) + " more";
        }

        private static string ContextLabel(Document doc, Element element)
        {
            var extras = new List<string>();
            var level = LevelName(doc, element);
            var view = OwnerViewName(doc, element);

            if (!string.IsNullOrWhiteSpace(level))
                extras.Add("Level: " + level);
            if (!string.IsNullOrWhiteSpace(view))
                extras.Add("View: " + view);

            var label = ElementLabel(element);
            return extras.Count == 0 ? label : label + " | " + string.Join(" | ", extras);
        }

        private static string CategoryName(Element element)
        {
            try
            {
                return element?.Category?.Name ?? element?.GetType().Name ?? "Element";
            }
            catch
            {
                return "Element";
            }
        }

        private static string ElementName(Element element)
        {
            try
            {
                if (!string.IsNullOrWhiteSpace(element?.Name))
                    return element.Name;
            }
            catch
            {
            }

            return "<unnamed>";
        }

        private static string ElementLabel(Element element)
        {
            return string.Format(
                "{0} {1} ({2})",
                CategoryName(element),
                ElementIdValue(element?.Id),
                ElementName(element));
        }

        private static string LevelName(Document doc, Element element)
        {
            if (doc == null || element == null)
                return "";

            foreach (var attribute in new[] { "LevelId", "GenLevel", "ReferenceLevel" })
            {
                try
                {
                    var property = element.GetType().GetProperty(attribute);
                    if (property == null)
                        continue;

                    var value = property.GetValue(element, null);
                    if (value is Level level && !string.IsNullOrWhiteSpace(level.Name))
                        return level.Name;

                    if (value is ElementId levelId && levelId != ElementId.InvalidElementId)
                    {
                        var docLevel = doc.GetElement(levelId) as Level;
                        if (docLevel != null && !string.IsNullOrWhiteSpace(docLevel.Name))
                            return docLevel.Name;
                    }
                }
                catch
                {
                }
            }

            return "";
        }

        private static string OwnerViewName(Document doc, Element element)
        {
            var ownerView = GetOwnerView(doc, element);
            return ownerView?.Name ?? "";
        }

        private static View GetOwnerView(Document doc, Element element)
        {
            if (doc == null || element == null)
                return null;

            try
            {
                if (element.OwnerViewId != ElementId.InvalidElementId)
                    return doc.GetElement(element.OwnerViewId) as View;
            }
            catch
            {
            }

            return null;
        }

        private static View GetSameLevelView(Document doc, Element element)
        {
            if (doc == null || element == null)
                return null;

            var targetLevel = LevelName(doc, element);
            if (string.IsNullOrWhiteSpace(targetLevel))
                return null;

            return new FilteredElementCollector(doc)
                .OfClass(typeof(View))
                .Cast<View>()
                .Where(IsNavigableView)
                .Where(view => string.Equals(LevelName(doc, view), targetLevel, StringComparison.OrdinalIgnoreCase))
                .OrderBy(view => view.ViewType.ToString(), StringComparer.OrdinalIgnoreCase)
                .ThenBy(view => view.Name, StringComparer.OrdinalIgnoreCase)
                .FirstOrDefault();
        }

        private static bool IsNavigableView(View view)
        {
            if (view == null || view.IsTemplate || view is ViewSheet)
                return false;

            try
            {
                if (!view.CanBePrinted)
                    return false;
            }
            catch
            {
                return false;
            }

            var rejected = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
            {
                "ProjectBrowser",
                "SystemBrowser",
                "Schedule",
                "DrawingSheet",
                "Internal",
                "Undefined",
                "Report",
                "CostReport",
                "LoadsReport",
                "ColumnSchedule",
                "PanelSchedule"
            };

            return !rejected.Contains(view.ViewType.ToString());
        }

        private static bool IsPropertyLine(Element element)
        {
            try
            {
                return string.Equals(element?.Category?.Name, "Property Lines", StringComparison.OrdinalIgnoreCase);
            }
            catch
            {
                return false;
            }
        }

        private static bool IsGroupMember(Element element)
        {
            try
            {
                return element != null && element.GroupId != ElementId.InvalidElementId;
            }
            catch
            {
                return false;
            }
        }


        private static Element GetGroupElement(Document doc, Element element)
        {
            try
            {
                if (doc == null || element == null || element.GroupId == ElementId.InvalidElementId)
                    return null;

                return doc.GetElement(element.GroupId);
            }
            catch
            {
                return null;
            }
        }

        private static Element TopParentElement(Document doc, Element primaryElement, Element contextElement, Element explicitGroup)
        {
            if (explicitGroup != null)
                return explicitGroup;

            foreach (var candidate in new[] { contextElement, primaryElement })
            {
                var group = GetGroupElement(doc, candidate);
                if (group != null)
                    return group;
            }

            return contextElement ?? primaryElement;
        }

        private static double? CurveAngleDegrees(Line line)
        {
            if (line == null)
                return null;

            var direction = line.Direction;
            if (Math.Abs(direction.X) < 1e-9 && Math.Abs(direction.Y) < 1e-9)
                return null;

            return NormalizeAngleDegrees(Math.Atan2(direction.Y, direction.X) * 180.0 / Math.PI);
        }

        private static Line RotateCurveAboutStart(Line line, double deltaRadians)
        {
            var start = line.GetEndPoint(0);
            var transform = Transform.CreateRotationAtPoint(XYZ.BasisZ, deltaRadians, start);
            return line.CreateTransformed(transform) as Line;
        }

        private static double NormalizeAngleDegrees(double angle)
        {
            var value = angle;
            while (value < 0.0)
                value += 360.0;
            while (value >= 360.0)
                value -= 360.0;
            return value;
        }

        private static double ShortestDeltaDegrees(double current, double target)
        {
            var delta = NormalizeAngleDegrees(target) - NormalizeAngleDegrees(current);
            while (delta <= -180.0)
                delta += 360.0;
            while (delta > 180.0)
                delta -= 360.0;
            return delta;
        }

        private static double DegreesToRadians(double degrees)
        {
            return degrees * Math.PI / 180.0;
        }

        private static string NormalizeText(string value)
        {
            return string.Join(" ", (value ?? "").Split((char[])null, StringSplitOptions.RemoveEmptyEntries)).Trim();
        }

        private static int ElementIdValue(ElementId elementId)
        {
#if NET48
            return elementId != null ? elementId.IntegerValue : -1;
#else
            return elementId != null ? (int)elementId.Value : -1;
#endif
        }

        private static Element GetElement(Document doc, int? elementIdValue)
        {
            if (doc == null || !elementIdValue.HasValue || elementIdValue.Value < 0)
                return null;

            return doc.GetElement(new ElementId(elementIdValue.Value));
        }

        private static void SelectElement(UIDocument uiDoc, int? elementIdValue)
        {
            if (uiDoc == null || !elementIdValue.HasValue)
                return;

            try
            {
                var ids = new Collection<ElementId> { new ElementId(elementIdValue.Value) };
                uiDoc.Selection.SetElementIds(ids);
            }
            catch
            {
            }
        }

        private int? ResolveEditableGroupId(UIDocument uiDoc)
        {
            if (uiDoc == null)
                return _activeGroupContextId;

            // First try: infer from current selection
            try
            {
                var selectedIds = uiDoc.Selection.GetElementIds();
                var groupIds = new HashSet<int>();
                foreach (var elementId in selectedIds)
                {
                    var element = uiDoc.Document.GetElement(elementId);
                    if (element == null || element.GroupId == ElementId.InvalidElementId)
                        continue;
                    groupIds.Add(ElementIdValue(element.GroupId));
                }

                if (groupIds.Count == 1)
                {
                    _activeGroupContextId = groupIds.First();
                    return _activeGroupContextId;
                }
            }
            catch
            {
            }

            // Second try: when in group edit mode the active view is scoped to the group
            // being edited. Collect all GroupIds from visible elements, then identify the
            // outermost group — the one whose own GroupId is InvalidElementId (i.e. not
            // itself nested inside another group). This handles groups with nested sub-groups.
            try
            {
                var view = uiDoc.ActiveView;
                if (view != null)
                {
                    var doc = uiDoc.Document;
                    var groupIds = new HashSet<int>();
                    var collector = new FilteredElementCollector(doc, view.Id)
                        .WhereElementIsNotElementType();
                    foreach (var element in collector)
                    {
                        if (element.GroupId != ElementId.InvalidElementId)
                            groupIds.Add(ElementIdValue(element.GroupId));
                    }

                    // Among the groups referenced by visible elements, prefer the one that
                    // is not itself a member of another group — that is the root editing group.
                    foreach (var gid in groupIds)
                    {
                        var groupElem = doc.GetElement(new ElementId(gid));
                        if (groupElem != null && groupElem.GroupId == ElementId.InvalidElementId)
                        {
                            _activeGroupContextId = gid;
                            return _activeGroupContextId;
                        }
                    }
                }
            }
            catch
            {
            }

            return _activeGroupContextId;
        }

        private static bool TryPostNamedCommand(UIApplication uiApp, params string[] postableNames)
        {
            foreach (var postableName in postableNames ?? Array.Empty<string>())
            {
                if (string.IsNullOrWhiteSpace(postableName))
                    continue;

                try
                {
                    var parsed = Enum.Parse(typeof(PostableCommand), postableName, true);
                    var commandId = RevitCommandId.LookupPostableCommandId((PostableCommand)parsed);
                    if (commandId != null && uiApp.CanPostCommand(commandId))
                    {
                        uiApp.PostCommand(commandId);
                        return true;
                    }
                }
                catch
                {
                }

                try
                {
                    var commandId = RevitCommandId.LookupCommandId(postableName);
                    if (commandId != null && uiApp.CanPostCommand(commandId))
                    {
                        uiApp.PostCommand(commandId);
                        return true;
                    }
                }
                catch
                {
                }
            }

            return false;
        }

        private static string WriteLog(Document doc, IList<RoundAnglesItem> applicable, int updated, int failed)
        {
            try
            {
                var folder = GetLogFolder(doc);
                Directory.CreateDirectory(folder);
                var fileName = "RoundOffAxisSketchLines_" + DateTime.Now.ToString("yyyyMMdd_HHmmss", CultureInfo.InvariantCulture) + ".log";
                var path = Path.Combine(folder, fileName);

                var builder = new StringBuilder();
                builder.AppendLine(Title + " Results");
                builder.AppendLine();
                builder.AppendLine("Updated: " + updated.ToString(CultureInfo.InvariantCulture));
                builder.AppendLine("Failed: " + failed.ToString(CultureInfo.InvariantCulture));
                builder.AppendLine();

                var updatedItems = applicable.Where(item => item.Status == RoundAnglesItemStatus.Updated).ToList();
                if (updatedItems.Count > 0)
                {
                    builder.AppendLine("Updated lines:");
                    foreach (var item in updatedItems)
                        builder.AppendLine("- " + item.DisplayText);
                    builder.AppendLine();
                }

                var failedItems = applicable.Where(item => item.Status == RoundAnglesItemStatus.Failed).ToList();
                if (failedItems.Count > 0)
                {
                    builder.AppendLine("Failures:");
                    foreach (var item in failedItems)
                        builder.AppendLine("- " + item.Label + " | " + item.ErrorMessage);
                }

                File.WriteAllText(path, builder.ToString());
                return path;
            }
            catch
            {
                return "";
            }
        }

        private static string GetLogFolder(Document doc)
        {
            try
            {
                if (doc != null && !string.IsNullOrWhiteSpace(doc.PathName))
                {
                    var folder = Path.GetDirectoryName(doc.PathName);
                    if (!string.IsNullOrWhiteSpace(folder))
                        return folder;
                }
            }
            catch
            {
            }

            return Path.GetTempPath();
        }
    }

    internal sealed class RoundAnglesExternalEventHandler : IExternalEventHandler
    {
        private readonly RoundAnglesController _controller;

        public RoundAnglesExternalEventHandler(RoundAnglesController controller)
        {
            _controller = controller;
        }

        public void Execute(UIApplication app)
        {
            _controller.Execute(app);
        }

        public string GetName()
        {
            return "WWPTools Round Off-Axis Sketch Lines";
        }
    }

    internal sealed class RoundAnglesRequest
    {
        public RoundAnglesRequest(RoundAnglesRequestKind kind)
        {
            Kind = kind;
        }

        public RoundAnglesRequestKind Kind { get; }

        public RoundAnglesItem Item { get; set; }

        public List<RoundAnglesItem> Items { get; set; }


    }

    internal enum RoundAnglesRequestKind
    {
        Refresh,
        Apply,
        Select,
        GoToView,
        EditGroup,
        EditBoundary,
        ZoomTo,
        Isolate
    }

    internal enum RoundAnglesItemStatus
    {
        Pending,
        Updated,
        Failed,
        GroupedSkip
    }

    internal sealed class RoundAnglesItem : INotifyPropertyChanged
    {
        private RoundAnglesItemStatus _status;
        private string _errorMessage;

        private RoundAnglesItem()
        {
        }

        public event PropertyChangedEventHandler PropertyChanged;

        public int WarningIndex { get; private set; }

        public int ElementIdValue { get; private set; }

        public int? ContextElementIdValue { get; private set; }

        public int? GroupElementIdValue { get; private set; }

        public int? PreferredUiElementIdValue
        {
            get
            {
                if (GroupElementIdValue.HasValue)
                    return GroupElementIdValue;
                if (ContextElementIdValue.HasValue)
                    return ContextElementIdValue;
                return ElementIdValue >= 0 ? ElementIdValue : (int?)null;
            }
        }

        public int? PreferredViewElementIdValue
        {
            get
            {
                if (ContextElementIdValue.HasValue)
                    return ContextElementIdValue;
                return PreferredUiElementIdValue;
            }
        }

        public string Label { get; private set; }

        public string ContextText { get; private set; }

        public string HostGroupText { get; private set; }

        public double CurrentAngle { get; private set; }

        public double TargetAngle { get; private set; }

        public double DeltaDegrees { get; private set; }

        public Line CurrentCurve { get; private set; }

        public Line RotatedCurve { get; private set; }

        public bool WasPinned { get; private set; }

        public bool IsGroupMember => Status == RoundAnglesItemStatus.GroupedSkip;

        public bool CanApply => Status == RoundAnglesItemStatus.Pending;

        public RoundAnglesItemStatus Status
        {
            get => _status;
            private set
            {
                if (_status == value)
                    return;

                _status = value;
                NotifyPresentationChanged();
            }
        }

        public string ErrorMessage
        {
            get => _errorMessage ?? "";
            private set
            {
                _errorMessage = value ?? "";
                NotifyPresentationChanged();
            }
        }

        public string DisplayText
        {
            get
            {
                var statusText = Status switch
                {
                    RoundAnglesItemStatus.Pending => "Pending",
                    RoundAnglesItemStatus.Updated => "Updated",
                    RoundAnglesItemStatus.Failed => "Failed",
                    RoundAnglesItemStatus.GroupedSkip => "Grouped Skip",
                    _ => "Item"
                };

                if (Status == RoundAnglesItemStatus.Failed && !string.IsNullOrWhiteSpace(ErrorMessage))
                    return string.Format("[{0}] {1} | {2}", statusText, CoreDisplayText, ErrorMessage);

                return string.Format("[{0}] {1}", statusText, CoreDisplayText);
            }
        }

        public string DetailText
        {
            get
            {
                var builder = new StringBuilder();
                builder.AppendLine("Element: " + Label);
                builder.AppendLine("Context: " + (string.IsNullOrWhiteSpace(ContextText) ? "(not available)" : ContextText));
                if (!string.IsNullOrWhiteSpace(HostGroupText))
                    builder.AppendLine("Host group: " + HostGroupText);

                if (Status != RoundAnglesItemStatus.GroupedSkip)
                {
                    builder.AppendLine("Current angle: " + CurrentAngle.ToString("F6", CultureInfo.InvariantCulture));
                    builder.AppendLine("Target angle: " + TargetAngle.ToString("F2", CultureInfo.InvariantCulture));
                    builder.AppendLine("Delta: " + DeltaDegrees.ToString("+0.000000;-0.000000;0.000000", CultureInfo.InvariantCulture) + " deg");
                }

                if (Status == RoundAnglesItemStatus.GroupedSkip)
                    builder.AppendLine("Reason: Group member. Open Edit Group, then refresh and apply inside that edit session.");

                if (Status == RoundAnglesItemStatus.Failed && !string.IsNullOrWhiteSpace(ErrorMessage))
                    builder.AppendLine("Error: " + ErrorMessage);

                return builder.ToString().Trim();
            }
        }

        private string CoreDisplayText
        {
            get
            {
                var parts = new List<string> { Label };
                if (!string.IsNullOrWhiteSpace(ContextText))
                    parts.Add("Context: " + ContextText);

                if (Status != RoundAnglesItemStatus.GroupedSkip)
                    parts.Add(string.Format(CultureInfo.InvariantCulture, "{0:F6} -> {1:F2} deg", CurrentAngle, TargetAngle));

                return string.Join(" | ", parts);
            }
        }

        public static RoundAnglesItem CreatePending(
            int warningIndex,
            Element element,
            Element contextElement,
            Element groupElement,
            string contextText,
            double currentAngle,
            double targetAngle,
            double deltaDegrees,
            Line currentCurve,
            Line rotatedCurve,
            bool wasPinned)
        {
            return new RoundAnglesItem
            {
                WarningIndex = warningIndex,
                ElementIdValue = GetId(element),
                ContextElementIdValue = GetNullableId(contextElement),
                GroupElementIdValue = GetNullableId(groupElement),
                Label = BuildLabel(element),
                ContextText = contextText ?? "",
                HostGroupText = groupElement != null ? BuildLabel(groupElement) : "",
                CurrentAngle = currentAngle,
                TargetAngle = targetAngle,
                DeltaDegrees = deltaDegrees,
                CurrentCurve = currentCurve,
                RotatedCurve = rotatedCurve,
                WasPinned = wasPinned,
                Status = RoundAnglesItemStatus.Pending,
                ErrorMessage = ""
            };
        }

        public static RoundAnglesItem CreateGroupedSkip(
            int warningIndex,
            Element element,
            Element contextElement,
            Element groupElement,
            string contextText)
        {
            return new RoundAnglesItem
            {
                WarningIndex = warningIndex,
                ElementIdValue = GetId(element),
                ContextElementIdValue = GetNullableId(contextElement),
                GroupElementIdValue = GetNullableId(groupElement),
                Label = BuildLabel(element),
                ContextText = contextText ?? "",
                HostGroupText = groupElement != null ? BuildLabel(groupElement) : "",
                Status = RoundAnglesItemStatus.GroupedSkip,
                ErrorMessage = ""
            };
        }

        public void MarkUpdated()
        {
            ErrorMessage = "";
            Status = RoundAnglesItemStatus.Updated;
        }

        public void MarkFailed(string errorMessage)
        {
            ErrorMessage = errorMessage;
            Status = RoundAnglesItemStatus.Failed;
        }

        private void NotifyPresentationChanged()
        {
            OnPropertyChanged(nameof(Status));
            OnPropertyChanged(nameof(IsGroupMember));
            OnPropertyChanged(nameof(CanApply));
            OnPropertyChanged(nameof(DisplayText));
            OnPropertyChanged(nameof(DetailText));
        }

        private void OnPropertyChanged([CallerMemberName] string propertyName = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }

        private static string BuildLabel(Element element)
        {
            if (element == null)
                return "Unknown element";

            var category = element.Category?.Name ?? element.GetType().Name;
            var name = !string.IsNullOrWhiteSpace(element.Name) ? element.Name : "<unnamed>";
            return string.Format("{0} {1} ({2})", category, GetId(element), name);
        }

        private static int GetId(Element element)
        {
            if (element == null)
                return -1;

#if NET48
            return element.Id.IntegerValue;
#else
            return (int)element.Id.Value;
#endif
        }

        private static int? GetNullableId(Element element)
        {
            return element != null ? GetId(element) : (int?)null;
        }
    }
}
