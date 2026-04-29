from __future__ import annotations

import math
from typing import Sequence

from geopy.geocoders import Nominatim
try:
    from shapely.geometry import Polygon
except Exception:  # pragma: no cover
    Polygon = None
try:
    import folium
    from streamlit_folium import st_folium
except Exception:  # pragma: no cover - fallback para ambiente sem libs opcionais
    folium = None
    st_folium = None


def geocode_address(address: str) -> tuple[float, float]:
    geolocator = Nominatim(user_agent="aurumcalc_map")
    location = geolocator.geocode(address, timeout=10)
    if location is None:
        raise ValueError("Endereço não encontrado")
    return float(location.latitude), float(location.longitude)


def calculate_polygon_area_m2(coordinates: Sequence[tuple[float, float]]) -> float:
    if len(coordinates) < 3:
        return 0.0

    lats = [c[0] for c in coordinates]
    lons = [c[1] for c in coordinates]
    lat0 = math.radians(sum(lats) / len(lats))
    r = 6378137.0
    xy = [(math.radians(lon) * r * math.cos(lat0), math.radians(lat) * r) for lat, lon in coordinates]
    if Polygon is not None:
        poly = Polygon(xy)
        return float(abs(poly.area))

    area = 0.0
    for i in range(len(xy)):
        x1, y1 = xy[i]
        x2, y2 = xy[(i + 1) % len(xy)]
        area += x1 * y2 - x2 * y1
    return abs(area) * 0.5


def render_area_map(center: tuple[float, float], zoom: int = 19, key: str = "area_map"):
    if folium is None or st_folium is None:
        raise RuntimeError("folium/streamlit-folium não disponíveis no ambiente.")
    fmap = folium.Map(location=center, zoom_start=zoom, tiles="OpenStreetMap")
    try:
        from folium.plugins import Draw

        Draw(
            export=True,
            draw_options={
                "polyline": True,
                "rectangle": True,
                "circle": False,
                "marker": False,
                "circlemarker": False,
                "polygon": True,
            },
            edit_options={"edit": True},
        ).add_to(fmap)
    except Exception:
        pass

    return st_folium(fmap, width="100%", height=520, key=key)


def calculate_azimuth_from_points(start_lat: float, start_lon: float, end_lat: float, end_lon: float) -> float:
    lat1 = math.radians(start_lat)
    lat2 = math.radians(end_lat)
    dlon = math.radians(end_lon - start_lon)
    x = math.sin(dlon) * math.cos(lat2)
    y = math.cos(lat1) * math.sin(lat2) - math.sin(lat1) * math.cos(lat2) * math.cos(dlon)
    bearing = math.degrees(math.atan2(x, y))
    return (bearing + 360.0) % 360.0
