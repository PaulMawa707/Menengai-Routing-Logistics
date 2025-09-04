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

# =============================
# Warehouses
# =============================
WAREHOUSES = {
    "MORL": {"lat": -0.28802969095623043, "lon": 36.04494759379902},
    "STO": {"lat": -1.3029821367535646, "lon": 36.865574991037754},
}

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
    if not isinstance(s, str):
        return ""
    return re.sub(r"[^A-Z0-9]", "", s.upper())


def extract_truck_number_from_text(text: str) -> str | None:
    if not isinstance(text, str):
        return None
    patterns = [
        r"Truck\s*(?:Number|No\.?|#)?\s*[:\-]?\s*([A-Z0-9\- ]{4,})",
        r"\b([A-Z]{2,3}\s*\d{3,4}\s*[A-Z])\b",
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



def read_asset_id_from_excel(excel_file, truck_number_norm):
    df = pd.read_excel(excel_file)
    df.columns = [col.strip().lower() for col in df.columns]
    name_col = None
    for candidate in ("reportname", "name", "unit", "unitname"):
        if candidate in df.columns:
            name_col = candidate
            break
    if name_col is None or "itemid" not in df.columns:
        raise ValueError("Assets Excel must contain columns like 'ReportName' (or 'Name') and 'itemId'.")
    df["normalized_name"] = df[name_col].astype(str).apply(normalize_plate)
    if not truck_number_norm:
        return None, None
    match = df[df["normalized_name"] == truck_number_norm]
    if match.empty:
        match = df[df["normalized_name"].str.contains(re.escape(truck_number_norm), na=False)]
    if not match.empty:
        row = match.iloc[0]
        return int(row["itemid"]), str(row[name_col])
    return None, None

# =============================
# Wialon API (Modified to support warehouse selection)
# =============================

def read_excel_to_df(excel_file):
    raw_df = pd.read_excel(excel_file, header=None)
    truck_number_norm = None
    if 0 in raw_df.columns:
        for row in raw_df[0].astype(str).tolist():
            truck_number_norm = extract_truck_number_from_text(row)
            if truck_number_norm:
                break

    # Force header row to row 8 (Excel row 8 = index 7)
    header_row_idx = 7

    df = pd.read_excel(excel_file, header=header_row_idx)
    df.columns = [
        re.sub(r"\s+", " ", str(col)).replace("\u00A0", " ").strip().upper()
        for col in df.columns
    ]
    df = df.loc[:, ~df.columns.str.startswith("UNNAMED")]

    required_cols = {"CUSTOMER ID", "CUSTOMER NAME", "LOCATION", "COORDINATES"}
    missing = required_cols - set(df.columns)
    if missing:
        raise ValueError(f"Missing required columns in orders Excel: {missing}")

    df = df[df["CUSTOMER ID"].notna()]
    df = df[~df["CUSTOMER NAME"].astype(str).str.contains("TOTAL", case=False, na=False)]

    for col in ("TONNAGE", "AMOUNT"):
        if col in df.columns:
            df[col] = pd.to_numeric(
                df[col].astype(str).str.replace(",", "").str.strip(), errors="coerce"
            ).fillna(0)
        else:
            df[col] = 0

    df_grouped = df.groupby(
        ["CUSTOMER ID", "CUSTOMER NAME", "LOCATION", "COORDINATES", "REP"],
        as_index=False,
    ).agg({
        "TONNAGE": "sum",
        "AMOUNT": "sum",
        "INVOICE NO.": lambda x: ", ".join(str(i) for i in x if pd.notna(i)) if "INVOICE NO." in df.columns else "",
    })

    df_grouped[["LAT", "LONG"]] = df_grouped["COORDINATES"].apply(
        lambda x: pd.Series(extract_coordinates(x))
    )
    df_grouped = df_grouped.dropna(subset=["LAT", "LONG"])
    return df_grouped, truck_number_norm

# PATCH: enforce nearest-first sequence regardless of Wialon optimizer
# Paste this function over your existing `send_orders_and_create_route`.

def send_orders_and_create_route(token, resource_id, unit_id, vehicle_name, df_grouped, tf, tt, warehouse_choice):
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

        # ---- Selected warehouse ----
        wh = WAREHOUSES[warehouse_choice]
        wh_lat, wh_lon = wh["lat"], wh["lon"]
        wh_name = warehouse_choice

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

        # Compute distance from warehouse and sort NEAREST first for sequence
        df_grouped['Distance_From_Warehouse'] = df_grouped.apply(
            lambda row: calculate_distance(wh_lat, wh_lon, row['LAT'], row['LONG']), axis=1
        )
        df_grouped = df_grouped.sort_values('Distance_From_Warehouse', ascending=True).reset_index(drop=True)

        # ---- Build orders for optimization (we will NOT use the returned sequence) ----
        orders = []
        wh_coords = f"{wh_lat}, {wh_lon}"

        for idx, row in df_grouped.iterrows():
            try:
                weight_kg = int(float(row.get('TONNAGE', 0)) * 1000)
            except Exception:
                weight_kg = 0
            coords = f"{row['LAT']}, {row['LONG']}"
            location = f"{row['LOCATION']} ({coords})"
            order_id = idx + 1  # we will use our own sequence later
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
                    "ut": 3600,
                    "rep": True,
                    "w": weight_kg,
                    "v": 0,
                    "pr": order_id,
                    "criterions": {"max_late": 0, "use_unloading_late": 0},
                    "a": location,
                },
                "cmp": {"unitRequirements": {"values": []}},
            })

        optimize_payload = {
            "svc": "order/optimize",
            "params": json.dumps({
                "itemId": int(resource_id),
                "orders": orders,
                "warehouses": [
                    {"id": 0, "y": wh_lat, "x": wh_lon, "n": wh_name, "f": 260, "a": f"{wh_name} ({wh_coords})"},
                    {"id": 99999, "y": wh_lat, "x": wh_lon, "n": wh_name, "f": 264, "a": f"{wh_name} ({wh_coords})"},
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
                "pf": {"n": wh_name, "y": wh_lat, "x": wh_lon, "a": f"{wh_name} ({wh_coords})"},
                "pt": {"n": wh_name, "y": wh_lat, "x": wh_lon, "a": f"{wh_name} ({wh_coords})"},
                "tf": tf,
                "tt": tt,
            }),
            "sid": session_id,
        }

        st.info("Optimizing route (for stats only)...")
        optimize_response = requests.post(base_url, data=optimize_payload, timeout=60)
        optimize_result = optimize_response.json()

        # We'll keep route summary/polylines if available, but ignore the optimizer's order
        route_summary = None
        end_warehouse_rp = None
        try:
            unit_key = str(unit_id)
            if isinstance(optimize_result, dict) and unit_key in optimize_result:
                unit_obj = optimize_result[unit_key]
                if unit_obj.get('routes'):
                    route_summary = unit_obj['routes'][0]
                # capture any final warehouse polyline if present
                if unit_obj.get('orders'):
                    for resp_order in reversed(unit_obj['orders']):
                        if isinstance(resp_order, dict) and resp_order.get('f') == 264:
                            end_warehouse_rp = resp_order.get('rp') or resp_order.get('p')
                            break
        except Exception:
            pass

        # ---- Route build (NEAREST-FIRST SEQUENCE) ----
        route_orders = []
        current_time = int(time.time())
        route_id = current_time
        last_visit_time = int(tf)
        sequence_index = 0

        # Start at warehouse (f:260)
        route_orders.append({
            "uid": int(unit_id),
            "id": 0,
            "n": wh_name,
            "p": {"ut": 0, "rep": True, "w": "0", "c": "0",
                  "r": {"vt": last_visit_time, "ndt": 60, "id": route_id, "i": sequence_index, "m": 0, "t": 0},
                  "u": int(unit_id), "a": f"{wh_name} ({wh_lat}, {wh_lon})",
                  "weight": "0", "cost": "0"},
            "f": 260,
            "tf": tf,
            "tt": tt,
            "r": 100,
            "y": wh_lat,
            "x": wh_lon,
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

        prev_coords = {'y': wh_lat, 'x': wh_lon}

        def calc_dist(a, b):
            from math import sin, cos, sqrt, atan2, radians
            R = 6371
            y1, x1, y2, x2 = map(radians, [a['y'], a['x'], b['y'], b['x']])
            dlat, dlon = y2 - y1, x2 - x1
            aa = sin(dlat/2)**2 + cos(y1)*cos(y2)*sin(dlon/2)**2
            return 2 * R * atan2(sqrt(aa), (1-aa)**0.5)

        # Build the sequence strictly by nearest-first order we computed above
        for idx, cust_row in df_grouped.iterrows():
            order_id = idx + 1
            order_name = cust_row['CUSTOMER NAME']
            coords = {'y': float(cust_row['LAT']), 'x': float(cust_row['LONG'])}

            weight_kg = int(float(cust_row.get('TONNAGE', 0)) * 1000)
            cost_val = float(cust_row.get('AMOUNT', 0.0))
            location = f"{cust_row['LOCATION']} ({coords['y']}, {coords['x']})"

            # simple time plan: +10 minutes per stop from last
            order_tm = max(last_visit_time + 3600, int(tf))

            mileage = int(calc_dist(prev_coords, coords) * 1000)

            # OSRM polyline for leg
            order_rp = None
            try:
                osrm_url = (
                    f"https://router.project-osrm.org/route/v1/driving/"
                    f"{prev_coords['x']},{prev_coords['y']};{coords['x']},{coords['y']}?overview=full&geometries=polyline"
                )
                osrm_json = requests.get(osrm_url, timeout=15).json()
                if isinstance(osrm_json, dict) and osrm_json.get('routes'):
                    order_rp = osrm_json['routes'][0].get('geometry')
                    #st.info(f"Using OSRM polyline for leg to order {order_id} (nearest-first).")
            except Exception:
                pass

            sequence_index += 1
            route_orders.append({
                "uid": int(unit_id),
                "id": order_id,
                "n": order_name,
                "p": {
                    "ut": 3600,
                    "rep": True,
                    "w": str(weight_kg),
                    "c": str(int(cost_val)),
                    "r": {"vt": order_tm, "ndt": 60, "id": route_id,
                          "i": sequence_index, "m": mileage, "t": 0},
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

        # Close at warehouse (f:264)
        def calc_dist_pts(a, b):
            from math import sin, cos, sqrt, atan2, radians
            R = 6371
            y1, x1, y2, x2 = map(radians, [a['y'], a['x'], b['y'], b['x']])
            dlat, dlon = y2 - y1, x2 - x1
            aa = sin(dlat/2)**2 + cos(y1)*cos(y2)*sin(dlon/2)**2
            return 2 * R * atan2(sqrt(aa), (1-aa)**0.5)
        mileage_back = int(calc_dist_pts(prev_coords, {'y': wh_lat, 'x': wh_lon}) * 1000)
        final_id = max([o.get("id", 0) for o in route_orders]) + 1
        sequence_index += 1

        if not end_warehouse_rp:
            try:
                osrm_url = (
                    f"https://router.project-osrm.org/route/v1/driving/"
                    f"{prev_coords['x']},{prev_coords['y']};{wh_lon},{wh_lat}?overview=full&geometries=polyline"
                )
                osrm_json = requests.get(osrm_url, timeout=15).json()
                if isinstance(osrm_json, dict) and osrm_json.get('routes'):
                    end_warehouse_rp = osrm_json['routes'][0].get('geometry')
                    #st.info("Using OSRM polyline for final leg to warehouse.")
            except Exception:
                pass

        route_orders.append({
            "uid": int(unit_id),
            "id": final_id,
            "n": wh_name,
            "p": {"ut": 0, "rep": True, "w": "0", "c": "0",
                  "r": {"vt": last_visit_time + 3600, "ndt": 60, "id": route_id,
                        "i": sequence_index, "m": mileage_back, "t": 0},
                  "u": int(unit_id), "a": f"{wh_name} ({wh_lat}, {wh_lon})",
                  "weight": "0", "cost": "0"},
            "f": 264,
            "tf": tf,
            "tt": tt,
            "r": 100,
            "y": wh_lat,
            "x": wh_lon,
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
                        "exp": 0,
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

        st.info("Creating final route (nearest-first)...")
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
        excel_files = st.file_uploader("Upload Excel File(s) - All must be for the same truck", type=["xls", "xlsx"], accept_multiple_files=True)
        assets_file = st.file_uploader("Upload Excel File (Assets)", type=["xls", "xlsx"])
        selected_date = st.date_input("Select Route Date")
        col1, col2 = st.columns(2)
        with col1:
            start_hour = st.slider("Route Start Hour", 0, 23, 6)
        with col2:
            end_hour = st.slider("Route End Hour", start_hour + 1, 23, 18)
        warehouse_choice = st.selectbox("Select Warehouse", list(WAREHOUSES.keys()))
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
                if not unit_id:
                    st.error(f"Could not find unit ID for truck (normalized): {truck_number_norm or 'UNKNOWN'}.")
                    return
                st.info("Summary of orders:")
                st.write(f"Delivery points: {len(gdf_joined)}")
                st.write(f"Tonnage: {gdf_joined['TONNAGE'].sum():.2f}")
                st.write(f"Amount: {gdf_joined['AMOUNT'].sum():.2f}")
                result = send_orders_and_create_route(token, int(resource_id), unit_id, vehicle_name, gdf_joined, tf, tt, warehouse_choice)
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
