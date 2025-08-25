import streamlit as st
import pandas as pd
import requests
import json
import time
from datetime import datetime
import pytz
import os
import base64
import re

# Optional libs you imported earlier but aren't used directly here
# import pdfplumber, numpy as np, sklearn, shapely, geopandas

# =============================
# Helpers: UI + environment
# =============================
os.environ['TZ'] = 'Africa/Nairobi'
try:
    time.tzset()
except Exception:
    pass

st.set_page_config(page_title="Wialon Logistics Uploader", layout="wide")


def get_base64_image(image_path):
    with open(image_path, "rb") as img_file:
        return base64.b64encode(img_file.read()).decode()


def set_background():
    try:
        background_image = get_base64_image("pexels-pixabay-236722.jpg")
        st.markdown(
            f"""
            <style>
            .stApp {{
                background-image: url("data:image/jpg;base64,{background_image}");
                background-size: cover;
                background-position: center;
                background-repeat: no-repeat;
            }}
            </style>
            """,
            unsafe_allow_html=True,
        )
    except Exception:
        # Background is optional
        pass


def show_logo_top_right(image_path, width=120):
    try:
        logo_base64 = get_base64_image(image_path)
        st.markdown(
            f"""
            <div style="display: flex; justify-content: space-between; align-items: center;">
                <div></div>
                <div style="margin-right: 1rem;">
                    <img src="data:image/png;base64,{logo_base64}" width="{width}">
                </div>
            </div>
            """,
            unsafe_allow_html=True,
        )
    except Exception:
        pass


set_background()
show_logo_top_right("CT-Logo.jpg", width=120)
st.markdown("<br>", unsafe_allow_html=True)

# =============================
# Normalization & parsing
# =============================

def normalize_plate(s: str) -> str:
    """Uppercase and strip spaces/hyphens for robust matching."""
    if not isinstance(s, str):
        return ""
    return re.sub(r"[^A-Z0-9]", "", s.upper())


def extract_truck_number_from_text(text: str) -> str | None:
    """
    Extracts a truck number from free text with many possible label variants.
    Matches patterns like:
      - "Truck Number: KBX123Z"
      - "Truck No. KBX 123Z"
      - "Truck: KBX-123Z"
    Returns the *normalized* form (no spaces/dashes, uppercased), e.g. "KBX123Z".
    """
    if not isinstance(text, str):
        return None

    # Try several flexible patterns
    patterns = [
        r"Truck\s*(?:Number|No\.?|#)?\s*[:\-]?\s*([A-Z0-9\- ]{4,})",
        r"\b([A-Z]{2,3}\s*\d{3,4}\s*[A-Z])\b",  # bare plates like KBX 123Z
    ]
    for pat in patterns:
        m = re.search(pat, text, re.IGNORECASE)
        if m:
            return normalize_plate(m.group(1))
    return None


# =============================
# Excel readers
# =============================

def extract_coordinates(coord_str):
    try:
        if isinstance(coord_str, str) and ("LAT:" in coord_str and "LONG:" in coord_str):
            parts = coord_str.split("LONG:")
            latitude = float(parts[0].replace("LAT:", "").strip().replace(" ", ""))
            longitude = float(parts[1].strip().replace(" ", ""))
            return latitude, longitude
    except Exception:
        pass
    return None, None


def read_excel_to_df(excel_file):
    raw_df = pd.read_excel(excel_file, header=None)

    # ---- Extract truck number (robust) ----
    truck_number_norm = None
    if 0 in raw_df.columns:
        for row in raw_df[0].astype(str).tolist():
            truck_number_norm = extract_truck_number_from_text(row)
            if truck_number_norm:
                break

    # ---- Find header row (looking for a cell that equals "NO.") ----
    header_row_idx = None
    for idx, row in raw_df.iterrows():
        if any(str(cell).strip().upper() == "NO." for cell in row):
            header_row_idx = idx
            break
    if header_row_idx is None:
        raise ValueError("Could not locate header row (cell 'NO.' not found).")

    df = pd.read_excel(excel_file, header=header_row_idx)

    # ---- Normalize headers ----
    df.columns = [
        re.sub(r"\s+", " ", str(col)).replace("\u00A0", " ").strip().upper()
        for col in df.columns
    ]
    df = df.loc[:, ~df.columns.str.startswith("UNNAMED")]

    # ---- Required columns sanity ----
    required_cols = {"CUSTOMER ID", "CUSTOMER NAME", "LOCATION", "LOCATION COORDINATES"}
    missing = required_cols - set(df.columns)
    if missing:
        raise ValueError(f"Missing required columns in orders Excel: {missing}")

    # ---- Filter out totals/empties ----
    df = df[df["CUSTOMER ID"].notna()]
    df = df[~df["CUSTOMER NAME"].astype(str).str.contains("TOTAL", case=False, na=False)]

    # ---- Numerics ----
    for col in ("TONNAGE", "AMOUNT"):
        if col in df.columns:
            df[col] = pd.to_numeric(
                df[col].astype(str).str.replace(",", "").str.strip(), errors="coerce"
            ).fillna(0)
        else:
            df[col] = 0

    # ---- Grouping ----
    df_grouped = df.groupby(
        ["CUSTOMER ID", "CUSTOMER NAME", "LOCATION", "LOCATION COORDINATES", "REP"],
        as_index=False,
    ).agg({
        "TONNAGE": "sum",
        "AMOUNT": "sum",
        "INVOICE NO": lambda x: ", ".join(str(i) for i in x if pd.notna(i)) if "INVOICE NO" in df.columns else "",
    })

    # ---- Coordinates ----
    df_grouped[["LAT", "LONG"]] = df_grouped["LOCATION COORDINATES"].apply(
        lambda x: pd.Series(extract_coordinates(x))
    )
    df_grouped = df_grouped.dropna(subset=["LAT", "LONG"])  # keep only rows with coords

    return df_grouped, truck_number_norm


def read_asset_id_from_excel(excel_file, truck_number_norm):
    df = pd.read_excel(excel_file)
    df.columns = [col.strip().lower() for col in df.columns]

    # Expect reportname + itemid, but be forgiving
    name_col = None
    for candidate in ("reportname", "name", "unit", "unitname"):
        if candidate in df.columns:
            name_col = candidate
            break
    if name_col is None or "itemid" not in df.columns:
        raise ValueError("Assets Excel must contain columns like 'ReportName' (or 'Name') and 'itemId'.")

    df["normalized_name"] = df[name_col].astype(str).apply(normalize_plate)

    # If extraction failed, try to salvage by using the only unique plate in orders file later
    if not truck_number_norm:
        return None, None

    # Try exact match first
    match = df[df["normalized_name"] == truck_number_norm]

    # If no exact match, try contains (helps when asset names have prefixes/suffixes)
    if match.empty:
        match = df[df["normalized_name"].str.contains(re.escape(truck_number_norm), na=False)]

    if not match.empty:
        row = match.iloc[0]
        return int(row["itemid"]), str(row[name_col])

    return None, None


# =============================
# Wialon API
# =============================

def send_orders_and_create_route(token, resource_id, unit_id, vehicle_name, df_grouped, tf, tt):
    try:
        base_url = "https://hst-api.wialon.com/wialon/ajax.html"

        # ---- Login ----
        st.info("Logging in with token...")
        login_payload = {
            "svc": "token/login",
            "params": json.dumps({"token": str(token).strip()}),
        }
        headers = {"Content-Type": "application/x-www-form-urlencoded"}
        login_response = requests.post(base_url, data=login_payload, headers=headers, timeout=30)
        login_result = login_response.json()

        if not isinstance(login_result, dict) or "eid" not in login_result:
            return {"error": 1, "message": f"Login failed: {login_result}"}
        session_id = login_result["eid"]

        # ---- MORL warehouse (consider moving to UI in future) ----
        morl_lat = -0.28802969095623043
        morl_lon = 36.04494759379902

        # ---- Distance helper ----
        def calculate_distance(lat1, lon1, lat2, lon2):
            from math import sin, cos, sqrt, atan2, radians
            R = 6371
            lat1, lon1, lat2, lon2 = map(radians, [lat1, lon1, lat2, lon2])
            dlat = lat2 - lat1
            dlon = lon2 - lon1
            a = sin(dlat / 2) ** 2 + cos(lat1) * cos(lat2) * sin(dlon / 2) ** 2
            c = 2 * atan2(sqrt(a), sqrt(1 - a))
            return R * c

        df_grouped['Distance_From_MORL'] = df_grouped.apply(
            lambda row: calculate_distance(morl_lat, morl_lon, row['LAT'], row['LONG']), axis=1
        )
        df_grouped = df_grouped.sort_values('Distance_From_MORL').reset_index(drop=True)

        # ---- Build orders for optimization ----
        orders = []
        morl_coords = f"{morl_lat}, {morl_lon}"
        order_id_to_name = {}

        for idx, row in df_grouped.iterrows():
            try:
                weight_kg = int(float(row.get('TONNAGE', 0)) * 1000)
            except Exception:
                weight_kg = 0
            coords = f"{row['LAT']}, {row['LONG']}"
            location = f"{row['LOCATION']} ({coords})"
            order_id = idx + 1
            orders.append({
                "y": float(row['LAT']),
                "x": float(row['LONG']),
                "tf": tf,
                "tt": tt,
                "n": row['CUSTOMER NAME'],
                "f": 0,
                "r": 20,
                "id": order_id,
                "p": {
                    "ut": 900,
                    "rep": True,
                    "w": weight_kg,
                    "v": 0,
                    "pr": idx + 1,
                    "criterions": {"max_late": 0, "use_unloading_late": 0},
                    "a": location,
                },
                "cmp": {"unitRequirements": {"values": []}},
            })
            order_id_to_name[order_id] = row['CUSTOMER NAME']

        optimize_payload = {
            "svc": "order/optimize",
            "params": json.dumps({
                "itemId": int(resource_id),
                "orders": orders,
                "warehouses": [
                    {"id": 0, "y": morl_lat, "x": morl_lon, "n": "MORL", "f": 260, "a": f"MORL ({morl_coords})"},
                    {"id": 99999, "y": morl_lat, "x": morl_lon, "n": "MORL", "f": 264, "a": f"MORL ({morl_coords})"},
                ],
                "flags": 524419,
                "units": [int(unit_id)],
                "gis": {
                    "addPoints": 1,
                    "provider": 2,
                    "speed": 0,
                    "cityJams": 1,
                    "countryJams": 1,
                    "mode": "driving",
                    "departure_time": 1,
                    "avoid": [],
                    "traffic_model": "best_guess",
                },
                "priority": {},
                "criterions": {"penalties_profile": "balanced"},
                "pf": {"n": "MORL", "y": morl_lat, "x": morl_lon, "a": f"MORL ({morl_coords})"},
                "pt": {"n": "MORL", "y": morl_lat, "x": morl_lon, "a": f"MORL ({morl_coords})"},
                "tf": tf,
                "tt": tt,
            }),
            "sid": session_id,
        }

        st.info("Optimizing route...")
        optimize_response = requests.post(base_url, data=optimize_payload, timeout=60)
        optimize_result = optimize_response.json()

        if isinstance(optimize_result, dict) and optimize_result.get('error', 0) != 0:
            return {"error": optimize_result.get('error', 1), "message": optimize_result.get('reason', 'Optimization failed')}

        # Extract optimized orders & final warehouse polyline if present
        optimized_orders = []
        route_summary = None
        end_warehouse_rp = None
        unit_key = str(unit_id)
        if isinstance(optimize_result, dict) and unit_key in optimize_result:
            unit_obj = optimize_result[unit_key]
            optimized_orders = unit_obj.get('orders', [])
            if unit_obj.get('routes'):
                route_summary = unit_obj['routes'][0]
            for resp_order in reversed(optimized_orders or []):
                if isinstance(resp_order, dict) and resp_order.get('f') == 264:
                    end_warehouse_rp = resp_order.get('rp') or resp_order.get('p')
                    break
        if not optimized_orders:
            return {"error": 1, "message": "No optimized orders in response"}

        # Build mapping of original coords
        coord_map = {
            row['CUSTOMER NAME']: {'y': float(row['LAT']), 'x': float(row['LONG'])}
            for _, row in df_grouped.iterrows()
        }
        coord_map['MORL'] = {'y': morl_lat, 'x': morl_lon}

        # ---- Route build ----
        route_orders = []
        current_time = int(time.time())
        route_id = current_time
        last_visit_time = int(tf)
        sequence_index = 0

        # Start at MORL (f:260)
        route_orders.append({
            "uid": int(unit_id),
            "id": 0,
            "n": "MORL",
            "p": {"ut": 0, "rep": True, "w": "0", "c": "0", "r": {"vt": last_visit_time, "ndt": 60, "id": route_id, "i": sequence_index, "m": 0, "t": 0}, "u": int(unit_id), "a": f"MORL ({morl_lat}, {morl_lon})", "weight": "0", "cost": "0"},
            "f": 260,
            "tf": tf,
            "tt": tt,
            "r": 100,
            "y": morl_lat,
            "x": morl_lon,
            "s": 0,
            "sf": 0,
            "trt": 0,
            "st": current_time,
            "cnm": 0,
            "ej": {},
            "cf": {},
            "cmp": {"unitRequirements": {"values": []}},
            "gfn": {"geofences": {}},
            "callMode": "create",
            "u": int(unit_id),
            "weight": "0",
            "cost": "0",
            "cargo": {"weight": "0", "cost": "0"},
        })

        prev_coords = {'y': morl_lat, 'x': morl_lon}

        # Add optimized customer orders
        for resp in optimized_orders:
            order_id = resp.get('id') if isinstance(resp, dict) else None
            if order_id is None or order_id not in order_id_to_name:
                continue
            order_name = order_id_to_name[order_id]
            coords = coord_map.get(order_name, {'y': morl_lat, 'x': morl_lon})

            cust_rows = df_grouped[df_grouped['CUSTOMER NAME'] == order_name]
            weight_kg = int(float(cust_rows.iloc[0]['TONNAGE']) * 1000) if not cust_rows.empty else 0
            cost_val = float(cust_rows.iloc[0]['AMOUNT']) if not cust_rows.empty else 0.0

            location = f"{order_name} ({coords['y']}, {coords['x']})"
            if not cust_rows.empty:
                location = f"{cust_rows.iloc[0]['LOCATION']} ({coords['y']}, {coords['x']})"

            order_tm = resp.get('tm') if isinstance(resp, dict) else None
            order_rp = (resp.get('rp') or resp.get('p')) if isinstance(resp, dict) else None

            # Normalize planned time
            if not isinstance(order_tm, int) or order_tm <= 0:
                order_tm = last_visit_time + 600
            else:
                order_tm = max(order_tm, last_visit_time + 60, int(tf))

            # Compute mileage for leg
            def calc_dist(a, b):
                from math import sin, cos, sqrt, atan2, radians
                R = 6371
                y1, x1, y2, x2 = map(radians, [a['y'], a['x'], b['y'], b['x']])
                dlat, dlon = y2 - y1, x2 - x1
                aa = sin(dlat/2)**2 + cos(y1)*cos(y2)*sin(dlon/2)**2
                return 2 * R * atan2(sqrt(aa), (1-aa)**0.5)
            mileage = int(calc_dist(prev_coords, coords) * 1000)

            # OSRM fallback for polyline if missing
            if not order_rp:
                try:
                    osrm_url = (
                        f"https://router.project-osrm.org/route/v1/driving/"
                        f"{prev_coords['x']},{prev_coords['y']};{coords['x']},{coords['y']}?overview=full&geometries=polyline"
                    )
                    osrm_json = requests.get(osrm_url, timeout=15).json()
                    if isinstance(osrm_json, dict) and osrm_json.get('routes'):
                        order_rp = osrm_json['routes'][0].get('geometry')
                        st.info(f"Using OSRM fallback polyline for leg to order {order_id}.")
                except Exception:
                    pass

            sequence_index += 1
            route_orders.append({
                "uid": int(unit_id),
                "id": order_id,
                "n": order_name,
                "p": {
                    "ut": 900,
                    "rep": True,
                    "w": str(weight_kg),
                    "c": str(int(cost_val)),
                    "r": {"vt": order_tm, "ndt": 60, "id": route_id, "i": sequence_index, "m": mileage, "t": 0},
                    "u": int(unit_id),
                    "a": location,
                    "weight": str(weight_kg),
                    "cost": str(int(cost_val)),
                },
                "f": 0,
                "tf": tf,
                "tt": tt,
                "r": 20,
                "y": coords['y'],
                "x": coords['x'],
                "s": 0,
                "sf": 0,
                "trt": 0,
                "st": current_time,
                "cnm": 0,
                **({"rp": order_rp} if order_rp else {}),
                "ej": {},
                "cf": {},
                "cmp": {"unitRequirements": {"values": []}},
                "gfn": {"geofences": {}},
                "callMode": "create",
                "u": int(unit_id),
                "weight": str(weight_kg),
                "cost": str(int(cost_val)),
                "cargo": {"weight": str(weight_kg), "cost": str(int(cost_val))},
            })

            prev_coords = coords
            last_visit_time = order_tm

        # Close at MORL (f:264)
        def calc_dist_pts(a, b):
            from math import sin, cos, sqrt, atan2, radians
            R = 6371
            y1, x1, y2, x2 = map(radians, [a['y'], a['x'], b['y'], b['x']])
            dlat, dlon = y2 - y1, x2 - x1
            aa = sin(dlat/2)**2 + cos(y1)*cos(y2)*sin(dlon/2)**2
            return 2 * R * atan2(sqrt(aa), (1-aa)**0.5)
        mileage_back = int(calc_dist_pts(prev_coords, {'y': morl_lat, 'x': morl_lon}) * 1000)
        final_id = max([o.get("id", 0) for o in route_orders]) + 1
        sequence_index += 1

        if not end_warehouse_rp:
            try:
                osrm_url = (
                    f"https://router.project-osrm.org/route/v1/driving/"
                    f"{prev_coords['x']},{prev_coords['y']};{morl_lon},{morl_lat}?overview=full&geometries=polyline"
                )
                osrm_json = requests.get(osrm_url, timeout=15).json()
                if isinstance(osrm_json, dict) and osrm_json.get('routes'):
                    end_warehouse_rp = osrm_json['routes'][0].get('geometry')
                    st.info("Using OSRM fallback polyline for final leg to warehouse.")
            except Exception:
                pass

        route_orders.append({
            "uid": int(unit_id),
            "id": final_id,
            "n": "MORL",
            "p": {"ut": 0, "rep": True, "w": "0", "c": "0", "r": {"vt": last_visit_time + 600, "ndt": 60, "id": route_id, "i": sequence_index, "m": mileage_back, "t": 0}, "u": int(unit_id), "a": f"MORL ({morl_lat}, {morl_lon})", "weight": "0", "cost": "0"},
            "f": 264,
            "tf": tf,
            "tt": tt,
            "r": 100,
            "y": morl_lat,
            "x": morl_lon,
            "s": 0,
            "sf": 0,
            "trt": 0,
            "st": current_time,
            "cnm": 0,
            **({"rp": end_warehouse_rp} if end_warehouse_rp else {}),
            "ej": {},
            "cf": {},
            "cmp": {"unitRequirements": {"values": []}},
            "gfn": {"geofences": {}},
            "callMode": "create",
            "u": int(unit_id),
            "weight": "0",
            "cost": "0",
            "cargo": {"weight": "0", "cost": "0"},
        })

        total_mileage = sum(order['p']['r']['m'] for order in route_orders)
        total_cost = sum(float(order['p']['c']) for order in route_orders if order['f'] == 0)
        total_weight = sum(int(order['p']['w']) for order in route_orders if order['f'] == 0)

        batch_payload = {
            "svc": "core/batch",
            "params": json.dumps({
                "params": [{
                    "svc": "order/route_update",
                    "params": {
                        "itemId": int(resource_id),
                        "orders": route_orders,
                        "uid": route_id,
                        "callMode": "create",
                        "exp": 3600,
                        "f": 0,
                        "n": f"{vehicle_name} - {datetime.now().strftime('%Y-%m-%d %H:%M')}",
                        "summary": {
                            "countOrders": len(route_orders),
                            "duration": route_summary.get('duration', 0) if isinstance(route_summary, dict) else 0,
                            "mileage": total_mileage,
                            "priceMileage": float(total_mileage) / 1000,
                            "priceTotal": total_cost,
                            "weight": total_weight,
                            "cost": total_cost,
                        },
                    },
                }],
                "flags": 0,
            }),
            "sid": session_id,
        }

        st.info("Creating final route...")
        route_response = requests.post(base_url, data=batch_payload, timeout=60)
        route_result = route_response.json()

        if isinstance(route_result, list):
            has_error = any(isinstance(item, dict) and item.get('error', 0) != 0 for item in route_result)
            if not has_error:
                planning_url = f"https://apps.wialon.com/logistics/?lang=en&sid={session_id}#/distrib/step3"
                return {"error": 0, "message": "Route created successfully", "planning_url": planning_url, "optimize_result": optimize_result, "route_result": route_result}
            error_item = next((item for item in route_result if isinstance(item, dict) and item.get('error', 0) != 0), None)
            return {"error": (error_item or {}).get('error', 1), "message": (error_item or {}).get('reason', 'Unknown error in batch response')}

        if isinstance(route_result, dict) and route_result.get("error", 1) == 0:
            planning_url = f"https://apps.wialon.com/logistics/?lang=en&sid={session_id}#/distrib/step3"
            return {"error": 0, "message": "Route created successfully", "planning_url": planning_url, "optimize_result": optimize_result, "route_result": route_result}

        return {"error": 1, "message": f"Unexpected or error response: {route_result}"}

    except Exception as e:
        st.error(f"An unexpected error occurred: {str(e)}")
        try:
            st.write("Error location (line):", e.__traceback__.tb_lineno)
        except Exception:
            pass
        return {"error": 1, "message": f"An unexpected error occurred: {str(e)}"}


# =============================
# Orchestration
# =============================

def process_multiple_excels(excel_files):
    all_gdfs = []
    truck_numbers = set()
    for excel_file in excel_files:
        gdf_joined, truck_number = read_excel_to_df(excel_file)
        if gdf_joined is not None and len(gdf_joined):
            all_gdfs.append(gdf_joined)
        if truck_number:
            truck_numbers.add(truck_number)

    if not all_gdfs:
        raise ValueError("No valid data found in any of the Excel files.")

    if len(truck_numbers) > 1:
        raise ValueError(f"Multiple truck numbers found (after normalization): {', '.join(sorted(truck_numbers))}")

    combined_gdf = pd.concat(all_gdfs, ignore_index=True)
    combined_gdf = combined_gdf.drop_duplicates(subset=["CUSTOMER ID", "LOCATION"], keep="first")
    sole_truck = next(iter(truck_numbers)) if truck_numbers else None
    return combined_gdf, sole_truck


def run_wialon_uploader():
    st.subheader("\U0001F4E6 Logistics Excel Orders Uploader (via Logistics API)")

    with st.form("upload_form"):
        excel_files = st.file_uploader(
            "Upload Excel File(s) - All must be for the same truck",
            type=["xls", "xlsx"],
            accept_multiple_files=True,
        )
        assets_file = st.file_uploader("Upload Excel File (Assets)", type=["xls", "xlsx"])
        selected_date = st.date_input("Select Route Date")
        col1, col2 = st.columns(2)
        with col1:
            start_hour = st.slider("Route Start Hour", 0, 23, 6)
        with col2:
            end_hour = st.slider("Route End Hour", start_hour + 1, 23, 18)
        token = st.text_input("Enter your Wialon Token", type="password")
        resource_id = st.text_input("Enter Wialon Resource ID")
        show_debug = st.checkbox("Show debug info", value=False)
        submit_btn = st.form_submit_button("Upload and Dispatch")

    if submit_btn:
        if not excel_files or not assets_file or not token or not resource_id:
            st.error("Please upload orders Excel, assets Excel, token, and resource ID.")
            return

        try:
            with st.spinner("Processing..."):
                tz = pytz.timezone('Africa/Nairobi')
                start_time = tz.localize(datetime.combine(selected_date, datetime.min.time().replace(hour=start_hour)))
                end_time = tz.localize(datetime.combine(selected_date, datetime.min.time().replace(hour=end_hour)))
                tf, tt = int(start_time.timestamp()), int(end_time.timestamp())

                gdf_joined, truck_number_norm = process_multiple_excels(excel_files)
                if gdf_joined is None or gdf_joined.empty:
                    st.error("No delivery rows with valid coordinates were found.")
                    return

                unit_id, vehicle_name = read_asset_id_from_excel(assets_file, truck_number_norm)

                if show_debug:
                    st.write("Extracted (normalized) truck number:", truck_number_norm)
                    try:
                        df_assets_debug = pd.read_excel(assets_file)
                        name_col = None
                        for candidate in ("reportname", "name", "unit", "unitname"):
                            if candidate in map(str.lower, df_assets_debug.columns):
                                name_col = candidate
                                break
                        # If mixed case columns, rebuild mapping
                        cols_map = {c.lower(): c for c in df_assets_debug.columns}
                        if name_col:
                            #st.write("Assets preview:")
                            temp = df_assets_debug[[cols_map.get(name_col, name_col)]].head(20).copy()
                            temp['normalized'] = temp[cols_map.get(name_col, name_col)].astype(str).apply(normalize_plate)
                            #st.dataframe(temp)
                    except Exception:
                        pass

                if not unit_id:
                    st.error(
                        f"Could not find unit ID for truck (normalized): {truck_number_norm or 'UNKNOWN'}.\n"
                        "Tip: ensure asset 'ReportName' contains the plate; hyphens/spaces don't matter."
                    )
                    return

                st.info("Summary of orders:")
                st.write(f"Delivery points: {len(gdf_joined)}")
                st.write(f"Tonnage: {gdf_joined['TONNAGE'].sum():.2f}")
                st.write(f"Amount: {gdf_joined['AMOUNT'].sum():.2f}")

                result = send_orders_and_create_route(token, int(resource_id), unit_id, vehicle_name, gdf_joined, tf, tt)

                if result.get("error") == 0:
                    st.success("✅ Route created successfully!")
                    st.markdown(f"[Open Wialon Logistics]({result['planning_url']})", unsafe_allow_html=True)
                    st.balloons()
                else:
                    st.error(f"❌ Failed: {result.get('message', 'Unknown error')}")
        except Exception as e:
            st.error(f"❌ Error: {str(e)}")


if __name__ == "__main__":
    run_wialon_uploader()
