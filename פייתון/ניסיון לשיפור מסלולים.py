import os
import json
import time
import pandas as pd
import requests
from datetime import datetime
from openpyxl import load_workbook
from difflib import get_close_matches
from dotenv import load_dotenv

# ==========================================
# הגדרות משתמש
# ==========================================
load_dotenv()  # טוען משתני סביבה מ-.env
GOOGLE_MAPS_API_KEY = os.getenv("GOOGLE_MAPS_API_KEY", "")
if not GOOGLE_MAPS_API_KEY:
    raise RuntimeError("GOOGLE_MAPS_API_KEY חסר - הוסף אותו ל-.env בתיקיית הפרוייקט")

# === נתיבים ===
BASE_FOLDER = r"C:\Users\דרורנדל\OneDrive - HASHOMER HACHADASH\שולחן העבודה\חילוץ נתוני הסעות"

OUTPUT_FOLDER = os.path.join(BASE_FOLDER, "דוחות")
CACHE_FILE = os.path.join(OUTPUT_FOLDER, "routes_cache.json")

# ==========================================
# אתחול
# ==========================================

# לוג שגיאות מיקומים
location_errors_log = []

# מעקב אחר כל המיקומים שזוהו בהצלחה
known_locations = set()


# ==========================================
# אתחול
# ==========================================

# טעינת מטמון
def load_cache():
    """טוען מטמון מסלולים קיימים"""
    if os.path.exists(CACHE_FILE):
        try:
            with open(CACHE_FILE, 'r', encoding='utf-8') as f:
                return json.load(f)
        except:
            return {}
    return {}


def save_cache(cache):
    """שומר מטמון מסלולים"""
    with open(CACHE_FILE, 'w', encoding='utf-8') as f:
        json.dump(cache, f, ensure_ascii=False, indent=2)


routes_cache = load_cache()


def find_similar_location(failed_location, known_locations, cutoff=0.6):
    """
    מחפש מיקום דומה במיקומים המוכרים
    מחזיר רשימה של מיקומים דומים
    """
    if not failed_location or not known_locations:
        return []

    # חיפוש התאמות דומות
    matches = get_close_matches(failed_location, known_locations, n=3, cutoff=cutoff)
    return matches


def log_location_error(file_name, row_num, failed_location, location_type, full_route, similar_locations=None):
    """
    רושם שגיאת מיקום ללוג
    """
    entry = {
        "קובץ": file_name,
        "שורה": row_num,
        "מיקום בעייתי": failed_location,
        "סוג": location_type,  # "מוצא", "תחנת ביניים", "יעד"
        "מסלול מלא": full_route,
        "הצעות למיקומים דומים": ", ".join(similar_locations) if similar_locations else "לא נמצאו מיקומים דומים"
    }
    location_errors_log.append(entry)


def make_excel_rtl(filepath):
    """הופך את ה-Excel ל-RTL"""
    try:
        wb = load_workbook(filepath)
        for sheet in wb.worksheets:
            sheet.sheet_view.rightToLeft = True
        wb.save(filepath)
    except Exception as e:
        print(f"אזהרה: לא ניתן להפוך לRTL: {str(e)}")


def build_route_key(origin, waypoints, destination):
    """בונה מפתח ייחודי למסלול עבור המטמון"""
    if not origin or not destination:
        return None

    waypoints_str = waypoints if waypoints else ""
    return f"{origin}|{waypoints_str}|{destination}"


def get_route_from_google(origin, waypoints, destination, file_name, row_num):
    """
    מקבל מסלול מ-Google Maps API
    מחזיר: (distance_km, duration_minutes, maps_url, error)
    """

    # בדיקת מטמון
    cache_key = build_route_key(origin, waypoints, destination)
    if cache_key and cache_key in routes_cache:
        cached = routes_cache[cache_key]
        # הוסף מיקומים למאגר המיקומים המוכרים
        known_locations.add(origin)
        known_locations.add(destination)
        if waypoints:
            for wp in waypoints.split(';'):
                known_locations.add(wp.strip())
        return cached['distance_km'], cached['duration_min'], cached['maps_url'], None

    # בניית URL ל-API
    base_url = "https://maps.googleapis.com/maps/api/directions/json"

    params = {
        'origin': origin,
        'destination': destination,
        'key': GOOGLE_MAPS_API_KEY,
        'language': 'iw',  # עברית
        'region': 'IL'  # ישראל
    }

    # בניית מסלול מלא לתיעוד
    full_route = origin
    waypoints_list = []
    if waypoints:
        waypoints_list = [w.strip() for w in waypoints.split(';') if w.strip()]
        if waypoints_list:
            params['waypoints'] = '|'.join(waypoints_list)
            full_route += " → " + " → ".join(waypoints_list)
    full_route += " → " + destination

    try:
        response = requests.get(base_url, params=params, timeout=10)
        response.raise_for_status()
        data = response.json()

        if data['status'] != 'OK':
            # ניתוח השגיאה - איזה מיקום בעייתי
            error_details = analyze_location_error(data, origin, waypoints_list, destination, file_name, row_num,
                                                   full_route)

            return None, None, None, error_details

        # הצלחה - הוסף את כל המיקומים למאגר המיקומים המוכרים
        known_locations.add(origin)
        known_locations.add(destination)
        for wp in waypoints_list:
            known_locations.add(wp)

        # חישוב סך המרחק והזמן
        route = data['routes'][0]
        legs = route['legs']

        total_distance_m = sum(leg['distance']['value'] for leg in legs)
        total_duration_s = sum(leg['duration']['value'] for leg in legs)

        distance_km = round(total_distance_m / 1000, 2)
        duration_min = round(total_duration_s / 60, 1)

        # בניית URL למפה
        maps_url = build_google_maps_url(origin, waypoints, destination)

        # שמירה במטמון
        if cache_key:
            routes_cache[cache_key] = {
                'distance_km': distance_km,
                'duration_min': duration_min,
                'maps_url': maps_url
            }
            save_cache(routes_cache)

        return distance_km, duration_min, maps_url, None

    except requests.exceptions.Timeout:
        return None, None, None, "Google Maps לא הגיב בזמן (timeout)"
    except requests.exceptions.RequestException as e:
        return None, None, None, f"בעיית תקשורת עם Google Maps"
    except Exception as e:
        return None, None, None, f"שגיאה לא צפויה בחישוב המסלול"


def analyze_location_error(api_response, origin, waypoints_list, destination, file_name, row_num, full_route):
    """
    מנתח את תגובת ה-API כדי לזהות איזה מיקום בעייתי
    """
    status = api_response['status']

    if status == 'NOT_FOUND':
        # אחד המיקומים לא נמצא - צריך לזהות איזה
        error_msg = "Google Maps לא מצא את אחד המיקומים במסלול"

        # ננסה לזהות איזה מיקום בעייתי על ידי בדיקה נפרדת
        problematic_locations = []

        # בדיקת מוצא
        if not verify_single_location(origin):
            similar = find_similar_location(origin, known_locations)
            log_location_error(file_name, row_num, origin, "מוצא", full_route, similar)
            problematic_locations.append(f"מוצא: {origin}")

        # בדיקת תחנות ביניים
        for wp in waypoints_list:
            if not verify_single_location(wp):
                similar = find_similar_location(wp, known_locations)
                log_location_error(file_name, row_num, wp, "תחנת ביניים", full_route, similar)
                problematic_locations.append(f"תחנה: {wp}")

        # בדיקת יעד
        if not verify_single_location(destination):
            similar = find_similar_location(destination, known_locations)
            log_location_error(file_name, row_num, destination, "יעד", full_route, similar)
            problematic_locations.append(f"יעד: {destination}")

        if problematic_locations:
            return error_msg + " - " + ", ".join(problematic_locations)
        else:
            return error_msg

    elif status == 'ZERO_RESULTS':
        return "לא נמצא מסלול אפשרי בין המיקומים"

    elif status == 'MAX_WAYPOINTS_EXCEEDED':
        return "יותר מדי תחנות ביניים במסלול"

    elif status == 'INVALID_REQUEST':
        return "בקשה לא תקינה - בדוק את שמות המיקומים"

    elif status == 'OVER_QUERY_LIMIT':
        return "חריגה ממכסת השימוש ב-Google Maps API"

    elif status == 'REQUEST_DENIED':
        return "הבקשה נדחתה - בדוק את ה-API key"

    else:
        return f"שגיאה לא מוכרת מ-Google Maps"


def verify_single_location(location):
    """
    בודק אם מיקום בודד תקין ב-Google Maps
    מחזיר True אם המיקום נמצא, False אם לא
    """
    try:
        url = "https://maps.googleapis.com/maps/api/geocode/json"
        params = {
            'address': location,
            'key': GOOGLE_MAPS_API_KEY,
            'region': 'IL'
        }

        response = requests.get(url, params=params, timeout=5)
        data = response.json()

        return data['status'] == 'OK'
    except:
        return True  # במקרה של שגיאה, נניח שהמיקום תקין כדי לא להוסיף רעש


def build_google_maps_url(origin, waypoints, destination):
    """בונה URL לפתיחת המסלול ב-Google Maps"""
    base_url = "https://www.google.com/maps/dir/"

    locations = [origin]

    if waypoints:
        waypoints_list = [w.strip() for w in waypoints.split(';') if w.strip()]
        locations.extend(waypoints_list)

    locations.append(destination)

    # קידוד URL
    from urllib.parse import quote
    encoded_locations = [quote(loc) for loc in locations]

    return base_url + '/'.join(encoded_locations)


def calculate_vat(row):
    """
    משלים מחירים עם מע"מ
    מחזיר: (price_per_trip_incl, total_price_incl)
    """

    # מחיר בודד
    if pd.notna(row['מחיר בודד לפני מע״מ']) and pd.isna(row['מחיר בודד כולל מע״מ']):
        price_per_incl = round(row['מחיר בודד לפני מע״מ'] * 1.18, 2)
    else:
        price_per_incl = row['מחיר בודד כולל מע״מ']

    # מחיר סופי
    if pd.notna(row['מחיר סופי לפני מע״מ']) and pd.isna(row['מחיר סופי כולל מע״מ']):
        total_incl = round(row['מחיר סופי לפני מע״מ'] * 1.18, 2)
    else:
        total_incl = row['מחיר סופי כולל מע״מ']

    return price_per_incl, total_incl


def process_routes_file(input_file):
    """
    מעבד קובץ נתונים גולמיים:
    1. משלים מע"מ
    2. מחשב מרחקים וזמנים
    """

    print(f"\n=== תהליך 2: עיבוד מסלולים ===")
    print(f"קורא קובץ: {os.path.basename(input_file)}\n")

    # קריאת הקובץ
    df = pd.read_excel(input_file)

    print(f"נמצאו {len(df)} שורות לעיבוד")

    # עמודות חדשות
    df['מרחק (ק״מ)'] = None
    df['זמן נסיעה (דקות)'] = None
    df['קישור Google Maps'] = None
    df['שגיאת מפות'] = None

    # סטטיסטיקות
    routes_processed = 0
    routes_cached = 0
    routes_errors = 0
    routes_skipped = 0
    api_calls = 0

    for idx, row in df.iterrows():
        origin = row['מוצא']
        waypoints = row['תחנות'] if pd.notna(row['תחנות']) else ""
        destination = row['יעד']

        # 1. השלמת מע"מ
        price_per_incl, total_incl = calculate_vat(row)
        df.at[idx, 'מחיר בודד כולל מע״מ'] = price_per_incl
        df.at[idx, 'מחיר סופי כולל מע״מ'] = total_incl

        # 2. בדיקה אם צריך לחשב מסלול
        # דלג על שורות ללא מוצא/יעד (תוספות מחיר)
        if pd.isna(origin) or pd.isna(destination):
            routes_skipped += 1
            continue

        # בדיקת מטמון
        cache_key = build_route_key(origin, waypoints, destination)
        if cache_key and cache_key in routes_cache:
            # מצאנו במטמון
            cached = routes_cache[cache_key]
            df.at[idx, 'מרחק (ק״מ)'] = cached['distance_km']
            df.at[idx, 'זמן נסיעה (דקות)'] = cached['duration_min']
            df.at[idx, 'קישור Google Maps'] = cached['maps_url']
            routes_cached += 1
            routes_processed += 1
        else:
            # קריאה ל-API
            print(f"[{idx + 1}/{len(df)}] מחשב מסלול: {origin} → {destination}", end="")
            if waypoints:
                print(f" (דרך: {waypoints})", end="")
            print()

            file_name = os.path.basename(input_file)
            distance, duration, maps_url, error = get_route_from_google(origin, waypoints, destination, file_name,
                                                                        idx + 1)

            if error:
                df.at[idx, 'שגיאת מפות'] = error
                routes_errors += 1
                print(f"  ⚠️ שגיאה: {error}")
            else:
                df.at[idx, 'מרחק (ק״מ)'] = distance
                df.at[idx, 'זמן נסיעה (דקות)'] = duration
                df.at[idx, 'קישור Google Maps'] = maps_url
                routes_processed += 1
                print(f"  ✓ {distance} ק״מ, {duration} דקות")

            api_calls += 1

            # המתנה קצרה בין קריאות API
            time.sleep(0.1)

    # שמירת הקובץ המעובד
    timestamp = datetime.now().strftime('%Y%m%d_%H%M')
    output_file = os.path.join(OUTPUT_FOLDER, f"נתונים_מעובדים_{timestamp}.xlsx")

    df.to_excel(output_file, index=False)
    make_excel_rtl(output_file)

    # שמירת לוג שגיאות מיקומים
    if location_errors_log:
        log_file = os.path.join(OUTPUT_FOLDER, f"לוג_מיקומים_בעייתיים_{timestamp}.xlsx")
        df_log = pd.DataFrame(location_errors_log)
        df_log.to_excel(log_file, index=False)
        make_excel_rtl(log_file)
        print(f"✓ לוג מיקומים בעייתיים נשמר: {log_file}")

    print(f"\n=== סיכום ===")
    print(f"✓ שורות מעובדות: {routes_processed}")
    print(f"  - מהמטמון: {routes_cached}")
    print(f"  - קריאות API חדשות: {api_calls}")
    print(f"⚠️ שגיאות: {routes_errors}")
    if location_errors_log:
        print(f"⚠️ מיקומים בעייתיים: {len(location_errors_log)} (ראה קובץ לוג)")
    print(f"⊘ דולגו (תוספות מחיר): {routes_skipped}")
    print(f"\n✓ קובץ נשמר: {output_file}")

    return output_file


def main():
    """נקודת כניסה ראשית"""

    # בדיקת API key
    if GOOGLE_MAPS_API_KEY == "YOUR_API_KEY_HERE":
        print("❌ שגיאה: נא להזין Google Maps API key בתחילת הקובץ")
        return

    # חיפוש קובץ הנתונים הגולמיים האחרון
    files = [f for f in os.listdir(OUTPUT_FOLDER) if f.startswith('נתונים_גולמיים_') and f.endswith('.xlsx')]

    if not files:
        print("❌ לא נמצא קובץ נתונים גולמיים בתיקייה")
        print(f"   חפש בתיקייה: {OUTPUT_FOLDER}")
        return

    # מיון לפי תאריך (האחרון ביותר)
    files.sort(reverse=True)
    latest_file = os.path.join(OUTPUT_FOLDER, files[0])

    print(f"נמצא קובץ: {files[0]}")

    # עיבוד
    process_routes_file(latest_file)

    print("\n=== תהליך עיבוד הסתיים בהצלחה! ===")


if __name__ == "__main__":
    main()