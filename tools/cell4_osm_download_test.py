#!/usr/bin/env python3
# FRO Cell4 MOTOR OSM road acquisition test
# Derived from Cell4 v65.85.7.58 OSM acquisition settings.
# Downloads the same MOTOR highway classes from OSM via Overpass failover.
from __future__ import annotations

import json
import os
import sys
import time
import hashlib
from collections import Counter
from datetime import datetime, timezone
from pathlib import Path

import networkx as nx
import osmnx as ox

OUT = Path(os.environ.get("FRO_OSM_OUT", "artifacts"))
OUT.mkdir(parents=True, exist_ok=True)

# 105 NLSC points extent (WGS84) plus Cell4's 7 km road-network buffer.
# Point extent: lon 120.984080..121.295193, lat 24.895563..25.014115.
WEST, SOUTH, EAST, NORTH = 120.9150, 24.8320, 121.3643, 25.0770
BBOX = (WEST, SOUTH, EAST, NORTH)

CUSTOM_FILTER = (
    '["highway"~"motorway|motorway_link|trunk|trunk_link|'
    'primary|primary_link|secondary|secondary_link|'
    'tertiary|tertiary_link|unclassified|residential|'
    'living_street|service|track|road"]'
    '["area"!~"yes"]'
)
USEFUL_TAGS = [
    "access", "vehicle", "motor_vehicle", "motorcar", "motorcycle",
    "access:conditional", "vehicle:conditional", "motor_vehicle:conditional",
    "motorcar:conditional", "motorcycle:conditional", "service", "surface",
    "tracktype", "smoothness", "width", "maxwidth", "narrow", "barrier",
    "oneway", "oneway:motor_vehicle", "oneway:motorcycle",
    "official_name", "alt_name", "loc_name", "short_name",
    "name:zh", "name:zh-Hant", "description",
]
ENDPOINTS = [
    "https://overpass-api.de/api",
    "https://overpass.private.coffee/api",
    "https://maps.mail.ru/osm/tools/overpass/api",
]

def main():
    print("FRO MOTOR OSM acquisition test", flush=True)
    print(f"OSMnx={ox.__version__} NetworkX={nx.__version__}", flush=True)
    print(f"BBOX_WGS84={BBOX}", flush=True)
    print(f"CUSTOM_FILTER={CUSTOM_FILTER}", flush=True)

    ox.settings.requests_timeout = 240
    ox.settings.overpass_rate_limit = True
    ox.settings.use_cache = False
    ox.settings.log_console = True
    ox.settings.useful_tags_way = list(dict.fromkeys(
        list(getattr(ox.settings, "useful_tags_way", [])) + USEFUL_TAGS
    ))

    errors = []
    graph = None
    source = None
    for idx, endpoint in enumerate(ENDPOINTS, 1):
        ox.settings.overpass_url = endpoint
        print(f"\n[{idx}/{len(ENDPOINTS)}] Overpass endpoint: {endpoint}", flush=True)
        try:
            started = time.time()
            graph = ox.graph.graph_from_bbox(
                BBOX,
                network_type="drive_service",
                custom_filter=CUSTOM_FILTER,
                simplify=True,
                retain_all=True,
                truncate_by_edge=True,
            )
            if graph.number_of_nodes() == 0 or graph.number_of_edges() == 0:
                raise RuntimeError("OSM returned an empty graph")
            source = endpoint
            print(
                f"DOWNLOAD_PASS nodes={graph.number_of_nodes()} "
                f"edges={graph.number_of_edges()} elapsed_sec={time.time()-started:.1f}",
                flush=True,
            )
            break
        except Exception as exc:
            msg = f"{endpoint}: {type(exc).__name__}: {exc}"
            errors.append(msg)
            print(f"ENDPOINT_FAIL {msg}", flush=True)
            graph = None
            if idx < len(ENDPOINTS):
                delay = min(2 * idx, 8)
                print(f"Failover cooldown {delay}s", flush=True)
                time.sleep(delay)

    if graph is None:
        raise RuntimeError("All Cell4 Overpass endpoints failed:\n" + "\n".join(errors))

    # Preserve original OSM tags and produce an auditable summary before export.
    highway = Counter()
    access = Counter()
    for _, _, data in graph.edges(data=True):
        values = data.get("highway", [])
        if not isinstance(values, (list, tuple, set)):
            values = [values]
        for value in values:
            highway[str(value)] += 1
        for key in ("motorcycle", "motor_vehicle", "vehicle", "access"):
            value = data.get(key)
            if value is not None:
                access[f"{key}={value}"] += 1

    undirected = nx.Graph(graph)
    components = sorted((len(c) for c in nx.connected_components(undirected)), reverse=True)
    graph_path = OUT / "osm_graph_MOTOR.graphml"
    ox.save_graphml(graph, filepath=graph_path)
    import zipfile
    zip_path = OUT / "osm_graph_MOTOR.graphml.zip"
    with zipfile.ZipFile(zip_path, "w", compression=zipfile.ZIP_DEFLATED, compresslevel=6) as zf:
        zf.write(graph_path, arcname=graph_path.name)
    graph_path.unlink()

    sha = hashlib.sha256(zip_path.read_bytes()).hexdigest()
    summary = {
        "project": "巡田最速適 FRO",
        "purpose": "Cell4 v65.85.7.58 MOTOR OSM road acquisition test",
        "created_at_utc": datetime.now(timezone.utc).isoformat(),
        "osmnx_version": ox.__version__,
        "networkx_version": nx.__version__,
        "source_endpoint": source,
        "bbox_wgs84_west_south_east_north": list(BBOX),
        "buffer_m": 7000,
        "custom_filter": CUSTOM_FILTER,
        "nodes": int(graph.number_of_nodes()),
        "directed_edges": int(graph.number_of_edges()),
        "undirected_connected_components": len(components),
        "largest_component_nodes": components[0] if components else 0,
        "highway_edge_tag_counts": dict(highway.most_common()),
        "access_tag_counts": dict(access.most_common()),
        "graphml_zip": zip_path.name,
        "graphml_zip_bytes": zip_path.stat().st_size,
        "graphml_zip_sha256": sha,
        "endpoint_failures": errors,
        "note": "This is the raw Cell4 MOTOR candidate graph. The final Cell4 motorcycle access-hierarchy filtering and routing optimization have not been run in this workflow.",
    }
    (OUT / "osm_download_summary.json").write_text(
        json.dumps(summary, ensure_ascii=False, indent=2), encoding="utf-8"
    )
    print("\n=== OSM_DOWNLOAD_SUMMARY ===", flush=True)
    print(json.dumps(summary, ensure_ascii=False, indent=2), flush=True)
    print("OSM_ARTIFACTS_READY=PASS", flush=True)

if __name__ == "__main__":
    main()
