#! python3
import clr
import json
import math
import os
import re
import sys
import traceback
import urllib.parse
import urllib.request

from System import Uri
from System.Collections.Generic import List
from System.IO import File, StringReader
from System.Windows import Visibility
from System.Windows.Interop import WindowInteropHelper
from System.Windows.Markup import XamlReader
from System.Windows.Media.Imaging import BitmapCacheOption, BitmapImage
from System.Xml import XmlReader

clr.AddReference("RevitAPI")
clr.AddReference("RevitAPIUI")

from Autodesk.Revit import DB, UI
from pyrevit.framework import EventHandler


TITLE = "Web Context Builder"
APP_ID = "WWPTools.WebContextBuilder"
DEFAULT_RADIUS_M = 500.0
MIN_RADIUS_M = 50.0
MAX_RADIUS_M = 2000.0
USER_AGENT = "WWPTools.WebContextBuilder/1.0"
OVERPASS_ENDPOINTS = [
    "https://overpass-api.de/api/interpreter",
    "https://overpass.kumi.systems/api/interpreter",
    "https://overpass.osm.ch/api/interpreter",
    "https://lz4.overpass-api.de/api/interpreter",
    "https://z.overpass-api.de/api/interpreter",
]
PEDESTRIAN_HIGHWAYS = (
    "footway",
    "path",
    "cycleway",
    "pedestrian",
    "steps",
    "track",
    "bridleway",
    "corridor",
)
MAX_FAILURES_IN_REPORT = 30
MAP_CLICK_SCHEME = "pyrevit-map://click"


script_dir = os.path.dirname(__file__)
lib_path = os.path.abspath(os.path.join(script_dir, "..", "..", "..", "lib"))
if lib_path not in sys.path:
    sys.path.append(lib_path)

import WWP_uiUtils as ui


def _get_doc():
    try:
        uidoc = __revit__.ActiveUIDocument
        if uidoc is None:
            return None
        return uidoc.Document
    except Exception:
        return None


def _set_owner(window):
    try:
        helper = WindowInteropHelper(window)
        helper.Owner = __revit__.ActiveUIDocument.Application.MainWindowHandle
    except Exception:
        pass


def _load_logo(image_control):
    try:
        logo_path = os.path.join(lib_path, "WWPtools-logo.png")
        if image_control is None or not os.path.isfile(logo_path):
            return
        bitmap = BitmapImage()
        bitmap.BeginInit()
        bitmap.UriSource = Uri(logo_path)
        bitmap.CacheOption = BitmapCacheOption.OnLoad
        bitmap.EndInit()
        image_control.Source = bitmap
    except Exception:
        pass


def _get_project_address(doc):
    project_info = getattr(doc, "ProjectInformation", None)
    if project_info is None:
        return ""

    for param_name in ("Project Address", "Address", "Site Address"):
        try:
            param = project_info.LookupParameter(param_name)
            if param:
                value = param.AsString()
                if value and value.strip():
                    return value.strip()
        except Exception:
            pass
    return ""


def _get_project_site_location(doc):
    try:
        site = doc.SiteLocation
    except Exception:
        site = None
    if site is None:
        return None

    try:
        lat_deg = math.degrees(site.Latitude)
        lon_deg = math.degrees(site.Longitude)
    except Exception:
        return None

    if abs(lat_deg) < 1e-9 and abs(lon_deg) < 1e-9:
        return None
    if abs(lat_deg) > 90.0 or abs(lon_deg) > 180.0:
        return None
    return lat_deg, lon_deg


def _parse_radius(text):
    try:
        value = float((text or "").strip())
    except Exception:
        return None
    if value < MIN_RADIUS_M or value > MAX_RADIUS_M:
        return None
    return value


def _reverse_geocode(lat, lon):
    url = (
        "https://nominatim.openstreetmap.org/reverse?"
        "format=jsonv2&zoom=18&addressdetails=1&lat={}&lon={}"
    ).format(lat, lon)
    data = _http_get_json(url, timeout_sec=20)
    display_name = (data or {}).get("display_name") or ""
    if display_name.strip():
        return display_name.strip()
    return "{:.6f}, {:.6f}".format(lat, lon)


def _build_map_html(center_lat=None, center_lon=None, radius_m=0.0, label="", zoom=16):
    has_location = center_lat is not None and center_lon is not None
    lat_text = "null" if not has_location else "{:.8f}".format(float(center_lat))
    lon_text = "null" if not has_location else "{:.8f}".format(float(center_lon))
    radius_text = "{:.2f}".format(max(0.0, float(radius_m or 0.0)))
    zoom_text = str(int(zoom if zoom is not None else (16 if has_location else 2)))
    label_text = json.dumps(label or "")
    return """<!doctype html>
<html>
<head>
  <meta charset="utf-8">
  <meta http-equiv="X-UA-Compatible" content="IE=Edge">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <link rel="stylesheet" href="https://unpkg.com/leaflet@1.7.1/dist/leaflet.css" />
  <style>
    html, body {{ height: 100%; margin: 0; padding: 0; overflow: hidden; background: #f2f5fa; }}
    #map {{ height: 100%; width: 100%; }}
    .map-note {{
      position: absolute;
      left: 12px;
      right: 12px;
      bottom: 12px;
      z-index: 9999;
      background: rgba(255, 255, 255, 0.94);
      border: 1px solid #d7dce6;
      border-radius: 6px;
      padding: 8px 10px;
      color: #5e6b80;
      font-family: Segoe UI, sans-serif;
      font-size: 12px;
      box-shadow: 0 2px 8px rgba(29, 39, 56, 0.12);
    }}
  </style>
  <script src="https://unpkg.com/leaflet@1.7.1/dist/leaflet.js"></script>
</head>
<body>
  <div id="map"></div>
  <div class="map-note">Click anywhere on the map to set the context center.</div>
  <script type="text/javascript">
    (function () {{
      var lat = {lat_text};
      var lon = {lon_text};
      var radius = {radius_text};
      var zoom = {zoom_text};
      var label = {label_text};

      var map = L.map('map', {{ zoomControl: true }});
      L.tileLayer('https://{{s}}.tile.openstreetmap.org/{{z}}/{{x}}/{{y}}.png', {{
        maxZoom: 19,
        attribution: '&copy; OpenStreetMap contributors'
      }}).addTo(map);

      if (lat === null || lon === null) {{
        map.setView([20.0, 0.0], 2);
      }} else {{
        map.setView([lat, lon], zoom);
        var marker = L.marker([lat, lon]).addTo(map);
        if (radius > 0) {{
          L.circle([lat, lon], {{
            radius: radius,
            color: '#2D6CDF',
            weight: 2,
            fillColor: '#2D6CDF',
            fillOpacity: 0.12
          }}).addTo(map);
        }}
        if (label) {{
          marker.bindPopup(label).openPopup();
        }}
      }}

      map.on('click', function (evt) {{
        var clickLat = evt.latlng.lat.toFixed(6);
        var clickLon = evt.latlng.lng.toFixed(6);
        window.location.href = '{scheme}?lat=' + encodeURIComponent(clickLat) + '&lon=' + encodeURIComponent(clickLon);
      }});
    }})();
  </script>
</body>
</html>""".format(
        lat_text=lat_text,
        lon_text=lon_text,
        radius_text=radius_text,
        zoom_text=zoom_text,
        label_text=label_text,
        scheme=MAP_CLICK_SCHEME,
    )


def _render_map(browser, center_lat=None, center_lon=None, radius_m=0.0, label="", zoom=16):
    if browser is None:
        return
    browser.NavigateToString(_build_map_html(center_lat, center_lon, radius_m, label, zoom))


def _show_dialog(doc):
    xaml_path = os.path.join(script_dir, "WebContextBuilderDialog.xaml")
    if not os.path.isfile(xaml_path):
        raise Exception("Missing dialog XAML: {}".format(xaml_path))

    xaml_text = File.ReadAllText(xaml_path)
    xaml_reader = XmlReader.Create(StringReader(xaml_text))
    window = XamlReader.Load(xaml_reader)
    _set_owner(window)

    address_text = window.FindName("AddressTextBox")
    radius_text = window.FindName("RadiusTextBox")
    validation_text = window.FindName("ValidationText")
    run_button = window.FindName("RunButton")
    cancel_button = window.FindName("CancelButton")
    locate_button = window.FindName("LocateButton")
    logo_image = window.FindName("LogoImage")
    map_browser = window.FindName("MapBrowser")
    map_hint_text = window.FindName("MapHintText")
    buildings_checkbox = window.FindName("BuildingsCheckBox")
    roads_checkbox = window.FindName("RoadsCheckBox")
    parks_checkbox = window.FindName("ParksCheckBox")
    water_checkbox = window.FindName("WaterCheckBox")

    _load_logo(logo_image)

    initial_address = _get_project_address(doc) or ""
    initial_site_location = _get_project_site_location(doc)
    if address_text is not None:
        address_text.Text = initial_address
    if radius_text is not None:
        radius_text.Text = str(int(DEFAULT_RADIUS_M))

    result = {"ok": False}
    map_state = {"location": None, "label": ""}

    def _set_validation(message):
        if message:
            validation_text.Text = message
            validation_text.Visibility = Visibility.Visible
        else:
            validation_text.Text = ""
            validation_text.Visibility = Visibility.Collapsed

    def _set_map_hint(message):
        if map_hint_text is not None:
            map_hint_text.Text = message or "Click map to choose the import center."

    def _get_valid_radius_or_default():
        parsed_radius = _parse_radius(radius_text.Text if radius_text is not None else "")
        if parsed_radius is None:
            return DEFAULT_RADIUS_M
        return parsed_radius

    def _set_map_location(lat, lon, label, zoom=16, update_address=False):
        map_state["location"] = (float(lat), float(lon))
        map_state["label"] = label or ""
        if update_address and address_text is not None:
            address_text.Text = label or "{:.6f}, {:.6f}".format(float(lat), float(lon))
        _render_map(map_browser, float(lat), float(lon), _get_valid_radius_or_default(), label or "", zoom)
        _set_map_hint("Map centered at {:.6f}, {:.6f}".format(float(lat), float(lon)))

    def _show_default_map():
        _render_map(map_browser, None, None, 0.0, "", 2)
        _set_map_hint("Click map to choose the import center.")

    def _locate_from_address_or_site():
        address_value = (address_text.Text or "").strip() if address_text is not None else ""
        if address_value:
            lat, lon = _geocode_address(address_value)
            _set_map_location(lat, lon, address_value, zoom=16, update_address=False)
            return

        if initial_site_location is None:
            raise Exception("Enter an address or set the project site location in Revit.")

        lat, lon = initial_site_location
        _set_map_location(
            lat,
            lon,
            "Project site location ({:.6f}, {:.6f})".format(lat, lon),
            zoom=15,
            update_address=False,
        )

    def _on_map_navigating(sender, args):
        try:
            uri = args.Uri
            absolute_uri = uri.AbsoluteUri if uri is not None else ""
            if not absolute_uri.lower().startswith(MAP_CLICK_SCHEME):
                return

            args.Cancel = True
            parsed = urllib.parse.urlparse(absolute_uri)
            query = urllib.parse.parse_qs(parsed.query)
            lat = float((query.get("lat") or [None])[0])
            lon = float((query.get("lon") or [None])[0])
            label = _reverse_geocode(lat, lon)
            _set_map_location(lat, lon, label, zoom=17, update_address=True)
            _set_validation("")
        except Exception as ex:
            try:
                args.Cancel = True
            except Exception:
                pass
            _set_validation("Map click failed: {}".format(str(ex)))

    def _on_locate(sender, args):
        try:
            _locate_from_address_or_site()
            _set_validation("")
        except Exception as ex:
            _set_validation("Unable to locate address: {}".format(str(ex)))

    def _on_run(sender, args):
        address = (address_text.Text or "").strip()
        radius = _parse_radius(radius_text.Text)
        layers = {
            "buildings": bool(buildings_checkbox.IsChecked) if buildings_checkbox is not None else True,
            "roads": bool(roads_checkbox.IsChecked) if roads_checkbox is not None else True,
            "parks": bool(parks_checkbox.IsChecked) if parks_checkbox is not None else True,
            "water": bool(water_checkbox.IsChecked) if water_checkbox is not None else True,
        }
        if radius is None:
            _set_validation("Radius must be between {} and {} meters.".format(int(MIN_RADIUS_M), int(MAX_RADIUS_M)))
            return
        if not any(layers.values()):
            _set_validation("Select at least one layer to import.")
            return
        if not address and _get_project_site_location(doc) is None:
            _set_validation("Enter an address or set the project site location in Revit.")
            return
        result["ok"] = True
        result["address"] = address
        result["radius_m"] = radius
        result["layers"] = layers
        result["location"] = map_state.get("location")
        result["location_label"] = map_state.get("label") or ""
        window.DialogResult = True
        window.Close()

    def _on_cancel(sender, args):
        window.DialogResult = False
        window.Close()

    run_button.Click += EventHandler(_on_run)
    cancel_button.Click += EventHandler(_on_cancel)
    if locate_button is not None:
        locate_button.Click += EventHandler(_on_locate)
    if map_browser is not None:
        map_browser.Navigating += _on_map_navigating

    try:
        if initial_address:
            _locate_from_address_or_site()
        elif initial_site_location is not None:
            lat, lon = initial_site_location
            _set_map_location(
                lat,
                lon,
                "Project site location ({:.6f}, {:.6f})".format(lat, lon),
                zoom=15,
                update_address=False,
            )
        else:
            _show_default_map()
    except Exception:
        _show_default_map()

    if window.ShowDialog() != True:
        return None
    return result if result.get("ok") else None


def _http_get_json(url, timeout_sec=20):
    request = urllib.request.Request(url, headers={"User-Agent": USER_AGENT})
    response = urllib.request.urlopen(request, timeout=timeout_sec)
    try:
        return json.loads(response.read().decode("utf-8"))
    finally:
        response.close()


def _http_post_text(url, body_text, timeout_sec=20):
    payload = urllib.parse.urlencode({"data": body_text}).encode("utf-8")
    request = urllib.request.Request(
        url,
        data=payload,
        headers={"User-Agent": USER_AGENT, "Content-Type": "application/x-www-form-urlencoded"},
    )
    response = urllib.request.urlopen(request, timeout=timeout_sec)
    try:
        return response.read().decode("utf-8")
    finally:
        response.close()


def _geocode_address(address):
    query = urllib.parse.quote(address)
    url = "https://nominatim.openstreetmap.org/search?format=json&limit=1&q={}".format(query)
    data = _http_get_json(url, timeout_sec=20)
    if not data:
        raise Exception("Address not found.")
    return float(data[0]["lat"]), float(data[0]["lon"])


def _calculate_square_bounds(center_lat, center_lon, radius_m):
    lat_rad = math.radians(center_lat)
    meters_per_degree_lat = 111320.0
    meters_per_degree_lon = max(1e-9, meters_per_degree_lat * math.cos(lat_rad))
    delta_lat = radius_m / meters_per_degree_lat
    delta_lon = radius_m / meters_per_degree_lon
    return (
        center_lat - delta_lat,
        center_lon - delta_lon,
        center_lat + delta_lat,
        center_lon + delta_lon,
    )


def _build_overpass_query(center_lat, center_lon, radius_m):
    return """
[out:json][timeout:90];
(
  way(around:{radius},{lat},{lon})["building"];
  relation(around:{radius},{lat},{lon})["building"];
  way(around:{radius},{lat},{lon})["building:part"];
  relation(around:{radius},{lat},{lon})["building:part"];
  way(around:{radius},{lat},{lon})["highway"];
  way(around:{radius},{lat},{lon})["leisure"="park"];
  relation(around:{radius},{lat},{lon})["leisure"="park"];
  way(around:{radius},{lat},{lon})["landuse"~"grass|recreation_ground|village_green|greenfield"];
  relation(around:{radius},{lat},{lon})["landuse"~"grass|recreation_ground|village_green|greenfield"];
  way(around:{radius},{lat},{lon})["natural"="water"];
  relation(around:{radius},{lat},{lon})["natural"="water"];
  way(around:{radius},{lat},{lon})["water"];
  relation(around:{radius},{lat},{lon})["water"];
  way(around:{radius},{lat},{lon})["landuse"~"reservoir|basin"];
  relation(around:{radius},{lat},{lon})["landuse"~"reservoir|basin"];
  way(around:{radius},{lat},{lon})["waterway"];
);
out body geom;
""".strip().format(
        lat=center_lat,
        lon=center_lon,
        radius=int(round(radius_m)),
    )


def _fetch_overpass_json(query):
    failures = []
    empty_data = None
    for endpoint in OVERPASS_ENDPOINTS:
        try:
            data = json.loads(_http_post_text(endpoint, query, timeout_sec=45))
            if data.get("elements"):
                data["_endpoint"] = endpoint
                return data
            if empty_data is None:
                data["_endpoint"] = endpoint
                empty_data = data
        except Exception as ex:
            failures.append("{} | {}".format(endpoint, str(ex)))

    if empty_data is not None:
        return empty_data

    raise Exception("All Overpass endpoints failed.\n{}".format("\n".join(failures[:MAX_FAILURES_IN_REPORT])))


def _point_key(point):
    return (round(point[0], 8), round(point[1], 8))


def _geometry_to_points(geometry):
    points = []
    for item in geometry or []:
        try:
            points.append((float(item["lat"]), float(item["lon"])))
        except Exception:
            continue
    return points


def _close_ring(points):
    clean = [point for point in (points or []) if point is not None]
    if len(clean) < 3:
        return clean
    if _point_key(clean[0]) != _point_key(clean[-1]):
        clean.append(clean[0])
    return clean


def _assemble_geometry_rings(members, role_name):
    segments = []
    for member in members or []:
        if (member.get("type") or "").lower() != "way":
            continue

        role = (member.get("role") or "").strip().lower()
        if role_name == "inner":
            if role != "inner":
                continue
        elif role == "inner":
            continue

        points = _geometry_to_points(member.get("geometry"))
        if len(points) < 2:
            continue
        segments.append(points)

    rings = []
    while segments:
        chain = list(segments.pop(0))
        progressed = True
        while progressed and len(chain) >= 2 and _point_key(chain[0]) != _point_key(chain[-1]):
            progressed = False
            for index, candidate in enumerate(segments):
                candidate_start = _point_key(candidate[0])
                candidate_end = _point_key(candidate[-1])
                chain_start = _point_key(chain[0])
                chain_end = _point_key(chain[-1])

                if chain_end == candidate_start:
                    chain.extend(candidate[1:])
                elif chain_end == candidate_end:
                    chain.extend(list(reversed(candidate[:-1])))
                elif chain_start == candidate_end:
                    chain = candidate[:-1] + chain
                elif chain_start == candidate_start:
                    chain = list(reversed(candidate[1:])) + chain
                else:
                    continue

                segments.pop(index)
                progressed = True
                break

        chain = _close_ring(chain)
        unique_points = {_point_key(point) for point in chain[:-1]} if len(chain) > 1 else set()
        if len(chain) >= 4 and len(unique_points) >= 3:
            rings.append(chain)

    return rings


def _point_in_ring_latlon(point, ring):
    x = point[1]
    y = point[0]
    inside = False
    for index in range(len(ring) - 1):
        x1 = ring[index][1]
        y1 = ring[index][0]
        x2 = ring[index + 1][1]
        y2 = ring[index + 1][0]
        intersects = (y1 > y) != (y2 > y)
        if not intersects:
            continue
        slope_x = ((x2 - x1) * (y - y1) / ((y2 - y1) or 1e-12)) + x1
        if x < slope_x:
            inside = not inside
    return inside


def _is_building_tag(tags):
    for key in ("building", "building:part"):
        value = (tags.get(key) or "").strip().lower()
        if value and value not in ("no", "false", "0"):
            return True
    return False


def _is_park_tag(tags):
    leisure = (tags.get("leisure") or "").strip().lower()
    landuse = (tags.get("landuse") or "").strip().lower()
    natural = (tags.get("natural") or "").strip().lower()
    return (
        leisure in ("park", "garden", "playground", "recreation_ground")
        or landuse in ("grass", "recreation_ground", "village_green", "greenfield")
        or natural == "grassland"
    )


def _is_water_tag(tags):
    natural = (tags.get("natural") or "").strip().lower()
    landuse = (tags.get("landuse") or "").strip().lower()
    water = (tags.get("water") or "").strip().lower()
    waterway = (tags.get("waterway") or "").strip().lower()
    return (
        natural == "water"
        or landuse in ("reservoir", "basin")
        or bool(water)
        or waterway == "riverbank"
    )


def _collect_elements(elements):
    ways = {}
    relations = []
    roads = []
    waterways = []

    for element in elements or []:
        element_type = (element.get("type") or "").lower()
        element_id = str(element.get("id"))
        tags = element.get("tags") or {}

        if element_type == "way":
            ways[element_id] = element

            highway = (tags.get("highway") or "").strip().lower()
            if highway and highway not in PEDESTRIAN_HIGHWAYS:
                roads.append(element)

            waterway = (tags.get("waterway") or "").strip().lower()
            if waterway and waterway not in ("dam", "weir", "dock"):
                waterways.append(element)
        elif element_type == "relation":
            relations.append(element)

    return ways, relations, roads, waterways


def _build_area_features(ways, relations, predicate):
    features = []
    relation_way_ids = set()

    for relation in relations or []:
        tags = relation.get("tags") or {}
        if not predicate(tags):
            continue

        members = relation.get("members") or []
        outer_rings = _assemble_geometry_rings(members, "outer")
        if not outer_rings:
            continue

        inner_rings = _assemble_geometry_rings(members, "inner")
        for member in members:
            if (member.get("type") or "").lower() == "way" and member.get("ref") is not None:
                relation_way_ids.add(str(member.get("ref")))

        for outer_index, outer_ring in enumerate(outer_rings):
            assigned_inners = []
            for inner_ring in inner_rings:
                if _point_in_ring_latlon(inner_ring[0], outer_ring):
                    assigned_inners.append(inner_ring)

            features.append(
                {
                    "id": "relation/{}:{}".format(relation.get("id"), outer_index + 1),
                    "tags": tags,
                    "outer": outer_ring,
                    "inners": assigned_inners,
                }
            )

    for way_id, way in (ways or {}).items():
        if way_id in relation_way_ids:
            continue

        tags = way.get("tags") or {}
        if not predicate(tags):
            continue

        ring = _close_ring(_geometry_to_points(way.get("geometry")))
        unique_points = {_point_key(point) for point in ring[:-1]} if len(ring) > 1 else set()
        if len(ring) < 4 or len(unique_points) < 3:
            continue

        features.append(
            {
                "id": "way/{}".format(way_id),
                "tags": tags,
                "outer": ring,
                "inners": [],
            }
        )

    return features


def _project_latlon(lat, lon, origin_lat, origin_lon):
    earth_radius_m = 6378137.0
    lat_rad = math.radians(lat)
    lon_rad = math.radians(lon)
    origin_lat_rad = math.radians(origin_lat)
    origin_lon_rad = math.radians(origin_lon)
    x_m = (lon_rad - origin_lon_rad) * math.cos(origin_lat_rad) * earth_radius_m
    y_m = (lat_rad - origin_lat_rad) * earth_radius_m
    return x_m, y_m


def _project_ring(ring, origin_lat, origin_lon):
    return [_project_latlon(point[0], point[1], origin_lat, origin_lon) for point in ring or []]


def _remove_duplicate_xy(points, tolerance_m=0.05):
    cleaned = []
    for point in points or []:
        if not cleaned:
            cleaned.append(point)
            continue

        prev_x, prev_y = cleaned[-1]
        curr_x, curr_y = point
        if math.hypot(curr_x - prev_x, curr_y - prev_y) > tolerance_m:
            cleaned.append(point)

    if len(cleaned) > 1:
        first_x, first_y = cleaned[0]
        last_x, last_y = cleaned[-1]
        if math.hypot(first_x - last_x, first_y - last_y) <= tolerance_m:
            cleaned[-1] = cleaned[0]

    return cleaned


def _signed_area_xy(points):
    area = 0.0
    for index in range(len(points) - 1):
        x1, y1 = points[index]
        x2, y2 = points[index + 1]
        area += (x1 * y2) - (x2 * y1)
    return area / 2.0


def _orient_ring(points, clockwise):
    area = _signed_area_xy(points)
    is_clockwise = area < 0
    if is_clockwise != clockwise:
        return list(reversed(points))
    return points


def _parse_numeric(raw_text):
    text = (raw_text or "").strip().replace(",", ".")
    match = re.search(r"[-+]?\d*\.?\d+", text)
    if not match:
        return None
    try:
        return float(match.group(0))
    except Exception:
        return None


def _parse_height_m(tags):
    for key in ("height", "building:height", "est_height"):
        raw_value = tags.get(key)
        if not raw_value:
            continue

        value = _parse_numeric(raw_value)
        if value is None or value <= 0:
            continue

        text = raw_value.lower()
        if "ft" in text or "feet" in text or "'" in text:
            return value * 0.3048
        if "mm" in text:
            return value / 1000.0
        if "cm" in text:
            return value / 100.0
        return value
    return None


def _parse_levels(tags):
    for key in ("building:levels", "levels"):
        value = _parse_numeric(tags.get(key))
        if value is not None and value > 0:
            return value
    return None


def _get_building_height_m(tags):
    explicit_height = _parse_height_m(tags)
    if explicit_height is not None:
        return explicit_height, "height"

    levels = _parse_levels(tags)
    if levels is not None:
        return ((levels - 1.0) * 2.95) + 3.25, "levels"

    return 12.0, "default"


def _road_width_m(tags):
    explicit = _parse_numeric(tags.get("width"))
    if explicit is not None and explicit > 0:
        return explicit

    highway = (tags.get("highway") or "").strip().lower()
    widths = {
        "motorway": 18.0,
        "trunk": 14.0,
        "primary": 12.0,
        "secondary": 10.0,
        "tertiary": 8.0,
        "residential": 7.2,
        "service": 6.0,
        "living_street": 5.0,
        "unclassified": 6.5,
    }
    return widths.get(highway, 7.0)


def _waterway_width_m(tags):
    explicit = _parse_numeric(tags.get("width"))
    if explicit is not None and explicit > 0:
        return explicit

    waterway = (tags.get("waterway") or "").strip().lower()
    widths = {
        "river": 12.0,
        "canal": 8.0,
        "stream": 4.0,
        "ditch": 2.5,
        "drain": 2.5,
    }
    return widths.get(waterway, 5.0)


def _normalize_vector(dx, dy):
    length = math.hypot(dx, dy)
    if length <= 1e-9:
        return None
    return dx / length, dy / length


def _line_intersection(point_a1, point_a2, point_b1, point_b2):
    x1, y1 = point_a1
    x2, y2 = point_a2
    x3, y3 = point_b1
    x4, y4 = point_b2
    denominator = ((x1 - x2) * (y3 - y4)) - ((y1 - y2) * (x3 - x4))
    if abs(denominator) <= 1e-9:
        return None

    det_a = (x1 * y2) - (y1 * x2)
    det_b = (x3 * y4) - (y3 * x4)
    px = ((det_a * (x3 - x4)) - ((x1 - x2) * det_b)) / denominator
    py = ((det_a * (y3 - y4)) - ((y1 - y2) * det_b)) / denominator
    return px, py


def _segment_normal(point_a, point_b, left_side):
    vector = _normalize_vector(point_b[0] - point_a[0], point_b[1] - point_a[1])
    if vector is None:
        return None
    ux, uy = vector
    if left_side:
        return -uy, ux
    return uy, -ux


def _offset_point(prev_point, point, next_point, half_width, left_side):
    prev_normal = _segment_normal(prev_point, point, left_side)
    next_normal = _segment_normal(point, next_point, left_side)

    if prev_normal is None and next_normal is None:
        return point
    if prev_normal is None:
        return point[0] + (next_normal[0] * half_width), point[1] + (next_normal[1] * half_width)
    if next_normal is None:
        return point[0] + (prev_normal[0] * half_width), point[1] + (prev_normal[1] * half_width)

    prev_start = (prev_point[0] + (prev_normal[0] * half_width), prev_point[1] + (prev_normal[1] * half_width))
    prev_end = (point[0] + (prev_normal[0] * half_width), point[1] + (prev_normal[1] * half_width))
    next_start = (point[0] + (next_normal[0] * half_width), point[1] + (next_normal[1] * half_width))
    next_end = (next_point[0] + (next_normal[0] * half_width), next_point[1] + (next_normal[1] * half_width))

    intersection = _line_intersection(prev_start, prev_end, next_start, next_end)
    if intersection is not None and math.hypot(intersection[0] - point[0], intersection[1] - point[1]) <= (half_width * 8.0):
        return intersection

    blended = _normalize_vector(prev_normal[0] + next_normal[0], prev_normal[1] + next_normal[1])
    if blended is not None:
        dot_value = max(0.2, min(1.0, (blended[0] * prev_normal[0]) + (blended[1] * prev_normal[1])))
        scale = half_width / dot_value
        return point[0] + (blended[0] * scale), point[1] + (blended[1] * scale)

    return point[0] + (prev_normal[0] * half_width), point[1] + (prev_normal[1] * half_width)


def _polyline_to_buffer_polygon(points, width_m):
    clean = _remove_duplicate_xy(points, tolerance_m=0.05)
    if len(clean) < 2:
        return None

    if clean[0] == clean[-1]:
        clean = clean[:-1]
    if len(clean) < 2:
        return None

    half_width = max(0.5, width_m / 2.0)
    left_points = []
    right_points = []

    for index, point in enumerate(clean):
        if index == 0:
            left_normal = _segment_normal(clean[0], clean[1], True)
            right_normal = _segment_normal(clean[0], clean[1], False)
            left_points.append((point[0] + (left_normal[0] * half_width), point[1] + (left_normal[1] * half_width)))
            right_points.append((point[0] + (right_normal[0] * half_width), point[1] + (right_normal[1] * half_width)))
        elif index == len(clean) - 1:
            left_normal = _segment_normal(clean[-2], clean[-1], True)
            right_normal = _segment_normal(clean[-2], clean[-1], False)
            left_points.append((point[0] + (left_normal[0] * half_width), point[1] + (left_normal[1] * half_width)))
            right_points.append((point[0] + (right_normal[0] * half_width), point[1] + (right_normal[1] * half_width)))
        else:
            left_points.append(_offset_point(clean[index - 1], point, clean[index + 1], half_width, True))
            right_points.append(_offset_point(clean[index - 1], point, clean[index + 1], half_width, False))

    polygon = left_points + list(reversed(right_points))
    polygon = _close_ring(_remove_duplicate_xy(polygon, tolerance_m=0.05))
    unique_points = {_point_key(point) for point in polygon[:-1]} if len(polygon) > 1 else set()
    if len(polygon) < 4 or len(unique_points) < 3:
        return None
    return polygon


def _meters_to_internal(value_m):
    try:
        return DB.UnitUtils.ConvertToInternalUnits(float(value_m), DB.UnitTypeId.Meters)
    except Exception:
        return DB.UnitUtils.ConvertToInternalUnits(float(value_m), DB.DisplayUnitType.DUT_METERS)


def _pick_base_level(doc):
    levels = list(DB.FilteredElementCollector(doc).OfClass(DB.Level).WhereElementIsNotElementType())
    levels.sort(key=lambda level: level.Elevation)
    return levels[0] if levels else None


def _get_floor_types(doc):
    floor_types = list(DB.FilteredElementCollector(doc).OfClass(DB.FloorType))
    floor_types.sort(key=lambda floor_type: floor_type.Name.lower())
    return floor_types


def _pick_floor_type(doc, keywords):
    floor_types = _get_floor_types(doc)
    if not floor_types:
        return None

    lowered_keywords = [keyword.lower() for keyword in (keywords or [])]
    for floor_type in floor_types:
        name = floor_type.Name.lower()
        if any(keyword in name for keyword in lowered_keywords):
            return floor_type

    for floor_type in floor_types:
        name = floor_type.Name.lower()
        if "foundation" not in name and "slab edge" not in name:
            return floor_type

    return floor_types[0]


def _xy_ring_to_curve_loop(points, clockwise, base_elevation):
    ring = _close_ring(_remove_duplicate_xy(points, tolerance_m=0.01))
    unique_points = {_point_key(point) for point in ring[:-1]} if len(ring) > 1 else set()
    if len(ring) < 4 or len(unique_points) < 3:
        return None

    ordered_points = _orient_ring(ring, clockwise=clockwise)
    curve_loop = DB.CurveLoop()
    for index in range(len(ordered_points) - 1):
        start_point = ordered_points[index]
        end_point = ordered_points[index + 1]
        start_xyz = DB.XYZ(_meters_to_internal(start_point[0]), _meters_to_internal(start_point[1]), base_elevation)
        end_xyz = DB.XYZ(_meters_to_internal(end_point[0]), _meters_to_internal(end_point[1]), base_elevation)
        if start_xyz.IsAlmostEqualTo(end_xyz):
            continue
        curve_loop.Append(DB.Line.CreateBound(start_xyz, end_xyz))

    return curve_loop if curve_loop.GetExactLength() > 0 else None


def _feature_to_curve_loops(feature, origin_lat, origin_lon, base_elevation):
    outer_points = _project_ring(feature.get("outer") or [], origin_lat, origin_lon)
    outer_loop = _xy_ring_to_curve_loop(outer_points, clockwise=False, base_elevation=base_elevation)
    if outer_loop is None:
        return None

    curve_loops = List[DB.CurveLoop]()
    curve_loops.Add(outer_loop)

    for inner_ring in feature.get("inners") or []:
        inner_points = _project_ring(inner_ring, origin_lat, origin_lon)
        inner_loop = _xy_ring_to_curve_loop(inner_points, clockwise=True, base_elevation=base_elevation)
        if inner_loop is not None:
            curve_loops.Add(inner_loop)

    return curve_loops


def _polygon_to_curve_loops(outer_points, inner_rings, base_elevation):
    outer_loop = _xy_ring_to_curve_loop(outer_points, clockwise=False, base_elevation=base_elevation)
    if outer_loop is None:
        return None

    curve_loops = List[DB.CurveLoop]()
    curve_loops.Add(outer_loop)

    for inner_points in inner_rings or []:
        inner_loop = _xy_ring_to_curve_loop(inner_points, clockwise=True, base_elevation=base_elevation)
        if inner_loop is not None:
            curve_loops.Add(inner_loop)

    return curve_loops


def _set_comment(element, text):
    try:
        parameter = element.get_Parameter(DB.BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS)
        if parameter and not parameter.IsReadOnly:
            parameter.Set(text)
            return
    except Exception:
        pass

    try:
        parameter = element.LookupParameter("Comments")
        if parameter and not parameter.IsReadOnly:
            parameter.Set(text)
    except Exception:
        pass


def _create_building_mass(doc, feature, origin_lat, origin_lon, base_elevation):
    height_m, height_source = _get_building_height_m(feature.get("tags") or {})
    if height_m <= 0:
        raise Exception("Building height resolved to a non-positive value.")

    curve_loops = _feature_to_curve_loops(feature, origin_lat, origin_lon, base_elevation)
    if curve_loops is None or curve_loops.Count == 0:
        raise Exception("Failed to build a valid building footprint.")

    solid = DB.GeometryCreationUtilities.CreateExtrusionGeometry(
        curve_loops,
        DB.XYZ.BasisZ,
        _meters_to_internal(height_m),
    )

    errors = []
    for built_in_category in (DB.BuiltInCategory.OST_Mass, DB.BuiltInCategory.OST_GenericModel):
        shape = None
        try:
            shape = DB.DirectShape.CreateElement(doc, DB.ElementId(built_in_category))
            shape.ApplicationId = APP_ID
            shape.ApplicationDataId = feature.get("id") or "building"

            geometry = List[DB.GeometryObject]()
            geometry.Add(solid)
            shape.SetShape(geometry)
            _set_comment(shape, "Web Context Builder | {} | {:.2f} m | {}".format(feature.get("id"), height_m, height_source))
            return shape
        except Exception as ex:
            errors.append("{}: {}".format(str(built_in_category), str(ex)))
            if shape is not None:
                try:
                    doc.Delete(shape.Id)
                except Exception:
                    pass

    raise Exception("; ".join(errors))


def _create_floor(doc, curve_loops, floor_type, level, comment_text):
    if curve_loops is None or curve_loops.Count == 0:
        raise Exception("No valid floor boundary loops were produced.")

    floor = None
    try:
        floor = DB.Floor.Create(doc, curve_loops, floor_type.Id, level.Id)
    except Exception:
        curve_array = DB.CurveArray()
        for curve in curve_loops[0]:
            curve_array.Append(curve)
        floor = doc.Create.NewFloor(curve_array, floor_type, level, False)

    if floor is None:
        raise Exception("Revit did not create a floor element.")

    _set_comment(floor, comment_text)

    try:
        offset_param = floor.get_Parameter(DB.BuiltInParameter.FLOOR_HEIGHTABOVELEVEL_PARAM)
        if offset_param and not offset_param.IsReadOnly:
            offset_param.Set(0.0)
    except Exception:
        pass

    return floor


def _summarize_data(address_label, radius_m, endpoint, selected_layers, building_features, roads, park_features, water_features, waterways):
    selected_layer_names = []
    if (selected_layers or {}).get("buildings"):
        selected_layer_names.append("Buildings")
    if (selected_layers or {}).get("roads"):
        selected_layer_names.append("Roads")
    if (selected_layers or {}).get("parks"):
        selected_layer_names.append("Parks")
    if (selected_layers or {}).get("water"):
        selected_layer_names.append("Water")

    lines = [
        "Location: {}".format(address_label),
        "Radius: {} m".format(int(round(radius_m))),
        "Overpass endpoint: {}".format(endpoint or "<unknown>"),
        "Selected layers: {}".format(", ".join(selected_layer_names) if selected_layer_names else "<none>"),
        "",
        "Buildings: {}".format(len(building_features or [])),
        "Road centerlines: {}".format(len(roads or [])),
        "Parks: {}".format(len(park_features or [])),
        "Water areas: {}".format(len(water_features or [])),
        "Waterways: {}".format(len(waterways or [])),
        "",
        "Buildings will be created as Mass-category DirectShapes.",
        "Roads, parks, and water will be created as Floors.",
    ]
    return "\n".join(lines)


def _summarize_results(created_counts, failures):
    lines = [
        "Buildings created: {}".format(created_counts.get("buildings", 0)),
        "Road floors created: {}".format(created_counts.get("roads", 0)),
        "Park floors created: {}".format(created_counts.get("parks", 0)),
        "Water floors created: {}".format(created_counts.get("water", 0)),
        "Failures: {}".format(len(failures)),
    ]

    if failures:
        lines.append("")
        lines.append("Failures (first {}):".format(MAX_FAILURES_IN_REPORT))
        for failure in failures[:MAX_FAILURES_IN_REPORT]:
            lines.append("{} | {}".format(failure["id"], failure["reason"]))

    return "\n".join(lines)


def _project_way_points(way, origin_lat, origin_lon):
    points = _geometry_to_points(way.get("geometry"))
    projected = [_project_latlon(point[0], point[1], origin_lat, origin_lon) for point in points]
    return _remove_duplicate_xy(projected, tolerance_m=0.05)


def _is_area_way(tags):
    area_value = (tags.get("area") or "").strip().lower()
    return area_value in ("yes", "1", "true") or bool(tags.get("area:highway"))


def main():
    doc = _get_doc()
    if doc is None:
        UI.TaskDialog.Show(TITLE, "No active Revit document found.")
        return

    user_inputs = _show_dialog(doc)
    if user_inputs is None:
        return

    address = user_inputs.get("address") or ""
    radius_m = user_inputs.get("radius_m") or DEFAULT_RADIUS_M
    selected_layers = user_inputs.get("layers") or {}
    selected_location = user_inputs.get("location")
    selected_location_label = user_inputs.get("location_label") or ""

    use_selected_location = selected_location is not None and (not address or address == selected_location_label)
    if use_selected_location:
        center_lat, center_lon = selected_location
        address_label = address or selected_location_label or "Selected map location ({:.6f}, {:.6f})".format(center_lat, center_lon)
    elif address:
        center_lat, center_lon = _geocode_address(address)
        address_label = address
    else:
        location = _get_project_site_location(doc)
        if location is None:
            UI.TaskDialog.Show(TITLE, "Enter an address or set the project site location in Revit.")
            return
        center_lat, center_lon = location
        address_label = "Project site location ({:.6f}, {:.6f})".format(center_lat, center_lon)

    query = _build_overpass_query(center_lat, center_lon, radius_m)
    overpass_data = _fetch_overpass_json(query)
    elements = overpass_data.get("elements") or []
    if not elements:
        UI.TaskDialog.Show(TITLE, "No OSM context data was returned for the selected location.")
        return

    ways, relations, roads, waterways = _collect_elements(elements)
    building_features = _build_area_features(ways, relations, _is_building_tag) if selected_layers.get("buildings") else []
    park_features = _build_area_features(ways, relations, _is_park_tag) if selected_layers.get("parks") else []
    water_features = _build_area_features(ways, relations, _is_water_tag) if selected_layers.get("water") else []
    roads = roads if selected_layers.get("roads") else []
    waterways = waterways if selected_layers.get("water") else []

    preview_text = _summarize_data(
        address_label,
        radius_m,
        overpass_data.get("_endpoint"),
        selected_layers,
        building_features,
        roads,
        park_features,
        water_features,
        waterways,
    )
    if not ui.uiUtils_show_text_report(
        "{} - Preview".format(TITLE),
        preview_text,
        ok_text="Run Import",
        cancel_text="Cancel",
        width=760,
        height=560,
    ):
        return

    base_level = _pick_base_level(doc)
    if base_level is None:
        UI.TaskDialog.Show(TITLE, "No Revit levels were found in the active model.")
        return

    road_floor_type = _pick_floor_type(doc, ["road", "street", "asphalt", "pavement", "concrete"]) if roads else None
    park_floor_type = _pick_floor_type(doc, ["park", "landscape", "grass", "green"]) if park_features else None
    water_floor_type = _pick_floor_type(doc, ["water", "pond", "pool", "generic"]) if (water_features or waterways) else None
    if (roads and road_floor_type is None) or (park_features and park_floor_type is None) or ((water_features or waterways) and water_floor_type is None):
        UI.TaskDialog.Show(TITLE, "No floor types were found in the active model.")
        return

    created_counts = {"buildings": 0, "roads": 0, "parks": 0, "water": 0}
    failures = []
    base_elevation = base_level.Elevation

    transaction = DB.Transaction(doc, TITLE)
    started = False
    try:
        transaction.Start()
        started = True

        for feature in building_features:
            try:
                _create_building_mass(doc, feature, center_lat, center_lon, base_elevation)
                created_counts["buildings"] += 1
            except Exception as ex:
                failures.append({"id": feature.get("id") or "building", "reason": str(ex)})

        for way in roads:
            way_id = "way/{}".format(way.get("id"))
            try:
                tags = way.get("tags") or {}
                projected = _project_way_points(way, center_lat, center_lon)
                if len(projected) < 2:
                    raise Exception("Road geometry is too short.")

                if _is_area_way(tags):
                    ring = _close_ring(projected)
                    curve_loops = _polygon_to_curve_loops(ring, [], base_elevation)
                else:
                    polygon = _polyline_to_buffer_polygon(projected, _road_width_m(tags))
                    curve_loops = _polygon_to_curve_loops(polygon, [], base_elevation) if polygon else None

                _create_floor(
                    doc,
                    curve_loops,
                    road_floor_type,
                    base_level,
                    "Web Context Builder | Road | {}".format(way_id),
                )
                created_counts["roads"] += 1
            except Exception as ex:
                failures.append({"id": way_id, "reason": str(ex)})

        for feature in park_features:
            try:
                curve_loops = _feature_to_curve_loops(feature, center_lat, center_lon, base_elevation)
                _create_floor(
                    doc,
                    curve_loops,
                    park_floor_type,
                    base_level,
                    "Web Context Builder | Park | {}".format(feature.get("id")),
                )
                created_counts["parks"] += 1
            except Exception as ex:
                failures.append({"id": feature.get("id") or "park", "reason": str(ex)})

        for feature in water_features:
            try:
                curve_loops = _feature_to_curve_loops(feature, center_lat, center_lon, base_elevation)
                _create_floor(
                    doc,
                    curve_loops,
                    water_floor_type,
                    base_level,
                    "Web Context Builder | Water Area | {}".format(feature.get("id")),
                )
                created_counts["water"] += 1
            except Exception as ex:
                failures.append({"id": feature.get("id") or "water", "reason": str(ex)})

        for way in waterways:
            way_id = "way/{}".format(way.get("id"))
            tags = way.get("tags") or {}
            if _is_water_tag(tags):
                continue
            try:
                projected = _project_way_points(way, center_lat, center_lon)
                polygon = _polyline_to_buffer_polygon(projected, _waterway_width_m(tags))
                curve_loops = _polygon_to_curve_loops(polygon, [], base_elevation) if polygon else None
                _create_floor(
                    doc,
                    curve_loops,
                    water_floor_type,
                    base_level,
                    "Web Context Builder | Waterway | {}".format(way_id),
                )
                created_counts["water"] += 1
            except Exception as ex:
                failures.append({"id": way_id, "reason": str(ex)})

        transaction.Commit()
    except Exception as ex:
        if started:
            try:
                transaction.RollBack()
            except Exception:
                pass
        UI.TaskDialog.Show(TITLE + " - Error", "{}\n\n{}".format(ex, traceback.format_exc()))
        return

    ui.uiUtils_show_text_report(
        "{} - Results".format(TITLE),
        _summarize_results(created_counts, failures),
        ok_text="Close",
        cancel_text=None,
        width=760,
        height=560,
    )


if __name__ == "__main__":
    try:
        main()
    except Exception as exc:
        UI.TaskDialog.Show(TITLE + " - Error", "{}\n\n{}".format(exc, traceback.format_exc()))
