# QuickOSM Urban Context

## Install the exporter

In Revit, open **WWPTools > Context > Context Builder > QuickOSM Urban** with **Shift-click**. Install to the default QGIS profile or choose another profile's `processing/scripts` folder. Refresh QGIS's Processing Scripts provider or restart QGIS. No additional Python packages are needed.

## Export in QGIS

1. Run QuickOSM's Urban preset. Run the preset separately from the exporter; the exporter reads the layers QuickOSM generated.
2. In the Processing Toolbox, search for **Export QuickOSM Urban for Revit** under **Scripts > WWPTools**.
3. Choose the styled `road` layer and, optionally, the `building` polygon layer (exported whole, never cropped; height from the optional **Building height column** you pick, else OSM `height`, else `building:levels` x 3 m, else a stable pseudo-random 6-8 m; stored as `height_m`; Revit imports them as Mass-category extrusions) and **Export area**. The default, **Full road layer (ignore screen zoom)**, exports the entire layer. Choose **Specified extent / canvas** only when you deliberately want a crop, then supply that extent.
4. Choose **Export CRS (local projected metres)**, **Road widths**, reference scale and context name. The initial reference scale comes from the canvas; it controls renderer rules and style-mode widths, independently of export area.
5. Save a JSON file. Re-exporting to the same valid package keeps its identity. A new filename starts a separate import.

Loading a GeoPackage does not necessarily restore the QuickOSM style. The road layer must use the Urban rule-based style, as it does after running the Urban preset. Your `Road+CSRS_UTM_17N.gpkg` records **EPSG:2958 (NAD83(CSRS) / UTM zone 17N)** and covers approximately 3.02 by 2.58 km before road buffering. It contains no saved layer style and no width/lanes fields; apply the Urban road style. Physical mode will use road-class estimates for this file; style mode uses the styled line width instead.

### Road width modes

- **Match QGIS style at reference scale** (default for visual context): converts the Urban colored road-body line width into a fixed polygon width. For example, 1 mm at 1:2,500 becomes 2.5 m; at 1:10,000 it becomes 10 m. Widths already expressed in map units remain in projected metres, subject to QGIS map-unit scale limits. This preserves the road-body appearance at the chosen scale rather than asserting surveyed dimensions. Changing zoom later does not resize Revit geometry.
- **Physical / estimated widths**: uses OSM `width`/`est_width`, then lane counts, then a road-class estimate. Reference scale still controls which style rules/categories are active, but does not determine the physical road width.

The package records the selected mode, scale, widths and their provenance. The Revit review dialog shows the mode before import. Style mode supports fixed SimpleLine road bodies; variable or complex road bodies are rejected with guidance. Thin casings, dashed overlays, joins/caps and label layout are simplified, so it is not a pixel-perfect copy of the QGIS map. Style widths must be between 0.01 and 500 m; choose a suitable reference scale.

### Coordinate systems

The source road layer must have a valid, correct CRS. The canvas CRS is not assumed to be the source CRS. Ordinary OSM source coordinates are longitude/latitude, but the exporter reads the actual layer metadata rather than assuming this. If the source CRS is missing, identify it in QGIS first; assigning a CRS is not reprojection.

Select the projected metric CRS used by the project (for example its applicable UTM or MTM system). Geometry is transformed into it before buffering. Geographic degrees, Web Mercator (EPSG:3857), and foot-based export CRSs are refused because their coordinate distances should not be used directly for these physical road widths. The package records source/export CRS and WKT, units, grid-north orientation and projected origin.

Re-exporting an existing package retains its CRS; changing it requires a new package file. On repeat import, a changed export extent/local origin does not move the existing geographic context: the first import's projected reference and Revit anchor are reused. Revit shared coordinates and site location are never changed. The importer does not infer a survey transformation from Revit's site latitude/longitude. For survey-aligned work, pick the corresponding origin and rotation deliberately.

## Import in Revit

1. Open a project and the target plan. Click **QuickOSM Urban** normally and choose the JSON.
2. In the road-group checklist, uncheck any groups you do not want. Each option shows its name, floor-surface count and colour code. Road-name text is filtered with the selected groups; updates remember the previous selection.
3. On the first import, choose internal origin or pick the corresponding origin point in the plan, then enter a counterclockwise rotation. Geometry is flat at the plan's level.
4. Review the CRS, projected origin and counts; choose **Import Context**. Roads use distinct native 10 mm floor types and materials for each renderer category; road names become native 2.5 mm TextNotes. Set the plan scale as appropriate for the site.
5. Import the same package again to update its owned geometry/text in the original plan. Unrelated elements and imports from other packages remain. The operation is one Revit undo step; failure rolls it back. If replacement would delete a user-added dimension/tag/hosted element referencing an imported floor, the update is stopped and rolled back; remove or rehost that reference first. Tool-created types/materials have separate ownership and do not overwrite unrelated assets with matching names.

The JSON is data, not a script. Package validation and native boundary checks run before replacement. Short edges and returning spikes below Revit's geometry tolerance are collapsed before the native checks; entire road surfaces are not silently dropped. Invalid/empty/cancelled exports preserve the previous package file. Import cancellation leaves the model unchanged. The RVT is not saved automatically.

## Scope and limitations

- QGIS 3.34+, Revit 2024+. Road-only Urban rule-based styles; buildings are simple footprint extrusions (no roofs); no terrain or elevations.
- Current active renderer rules, checked rule visibility and chosen style scale determine exported roads/categories. Unmatched/non-rendered roads are omitted, matching the style's visibility.
- Solid body colours map to materials. Borders become thin floor outlines; dotted/dashed decorative centre overlays are not modeled as extra floors.
- Physical mode uses OSM `width`/`est_width`, including explicit feet conversion, then lanes or road class. Style mode explicitly converts symbol thickness at the reference scale into visual context geometry.
- Intersections are partitioned by symbol rendering pass to avoid overlapping flat floors. Bridges/tunnels are flattened, with their source tags retained in the package.
- Road-name text and rotations are preserved as straight labels derived from geometry. QGIS curved text, halos and exact collision placement are not reproduced. Long/dense names may need manual adjustment in Revit.
- Large exports are bounded (20,000 source features, 10,000 surfaces, 250,000 polygon vertices, 5,000 labels, 64 MB). Export smaller areas if limits or Revit boundary tolerance are exceeded. Near-axis sketch warnings can occur with GIS geometry.

Data attribution: OpenStreetMap contributors, ODbL; <https://www.openstreetmap.org/copyright>.
