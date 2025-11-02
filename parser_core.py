from pathlib import Path
from lxml import etree

# ---- Params to extract (Casualties removed as requested) ----
PARAM_MAP = {
    "History": "history",
    "EventDate": "event_date_reported",
    "EventTime": "event_time_reported",
    "District": "district",
    "State": "state",
    "length": "length_m",
    "Breadth": "breadth_m",
    "Height": "height_m",
    "TypeLandslide": "type_landslide",
    "Material": "material",
    "Occurrence": "occurrence",
    "StructureAffected": "structure",
    "TriggerLandslide": "trigger",
    "Causes": "causes",
    "LandslideCategory": "landslide_category",
    "Remedial": "remedial",
}

GPS_XP = {
    "featurecoords": "./gpsdetails/featurecoords/text()",
    "accuracy": "./gpsdetails/accuracy/text()",
    "altitude": "./gpsdetails/altitude/text()",
    "speed": "./gpsdetails/speed/text()",
    "timestamp": "./gpsdetails/timestamp/text()",
    "typegps": "./gpsdetails/typegps/text()",
}

OBS_SIMPLE_XP = {
    "seqno": "./seqno/text()",
    "featuretype": "./featuretype/text()",
}

PHOTO_FIELDS = ["photoname", "photolat", "photolon", "photoacc", "photodir"]

def _text(node, xp):
    try:
        res = node.xpath(xp)
        if isinstance(res, list):
            return str(res[0]).strip() if res else ""
        return str(res).strip()
    except Exception:
        return ""

def _param_value(obs, name):
    try:
        v = obs.xpath(f"./params/param[paramname='{name}']/paramvalue/text()")
        return str(v[0]).strip() if v else ""
    except Exception:
        return ""

def _split_featurecoords(s):
    if not s:
        return "", ""
    parts = s.split()
    if len(parts) != 2:
        return "", ""
    lon, lat = parts[0], parts[1]
    return lat.strip(), lon.strip()

def _photos_list(obs):
    lst = []
    nodes = obs.xpath("./photos/photo")
    for i, p in enumerate(nodes, start=1):
        item = {"index": i}
        for f in PHOTO_FIELDS:
            item[f] = _text(p, f"./{f}/text()")
        lst.append(item)
    return lst

def parse_xml_file(path):
    """Return (rows, errors). rows[i]['photos'] is a list of photo dicts (kept internal)."""
    rows, errors = [], []
    try:
        tree = etree.parse(path)
    except Exception as e:
        return [], [f"{path}: XML parse error → {e}"]

    root = tree.getroot()
    observations = root.xpath("//observations/observation")
    if not observations:
        return [], [f"{path}: no <observation> nodes found"]

    observer = _text(root, "./projectdetails/observername/text()")
    source_name = Path(path).name

    for obs in observations:
        row = {"source_file": source_name}

        # simple
        for k, xp in OBS_SIMPLE_XP.items():
            row[k] = _text(obs, xp)

        # gps
        gps = {k: _text(obs, xp) for k, xp in GPS_XP.items()}
        row["featurecoords_raw"] = gps.get("featurecoords", "")
        lat, lon = _split_featurecoords(row["featurecoords_raw"])
        row["lat"] = lat
        row["lon"] = lon
        row["altitude_m"] = gps.get("altitude", "")
        row["gps_accuracy"] = gps.get("accuracy", "")
        row["gps_speed"] = gps.get("speed", "")
        row["event_timestamp"] = gps.get("timestamp", "")
        row["gps_type"] = gps.get("typegps", "")

        # observer
        row["observer"] = observer

        # params (without Casualties)
        for p, col in PARAM_MAP.items():
            row[col] = _param_value(obs, p)

        # photos (kept internal; not shown unless needed by dropdown)
        row["photos"] = _photos_list(obs)

        # trim
        for k, v in list(row.items()):
            if isinstance(v, str):
                row[k] = v.strip()

        rows.append(row)

    return rows, errors
