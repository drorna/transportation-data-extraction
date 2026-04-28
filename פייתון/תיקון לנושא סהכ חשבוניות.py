import os
import json
import time
import pandas as pd
from google import genai
from google.genai import types
from datetime import datetime
from openpyxl import load_workbook
from dotenv import load_dotenv

# ==========================================
# הגדרות משתמש
# ==========================================
load_dotenv()  # טוען משתני סביבה מ-.env
GEMINI_API_KEY = os.getenv("GEMINI_API_KEY", "")
if not GEMINI_API_KEY:
    raise RuntimeError("GEMINI_API_KEY חסר - הוסף אותו ל-.env בתיקיית הפרוייקט")

# === נתיבים ===
BASE_FOLDER = r"C:\Users\דרורנדל\OneDrive - HASHOMER HACHADASH\שולחן העבודה\חילוץ נתוני הסעות"

INPUT_FOLDER = os.path.join(BASE_FOLDER, "חשבוניות")
OUTPUT_FOLDER = os.path.join(BASE_FOLDER, "דוחות")

# ==========================================
# אתחול
# ==========================================
client = genai.Client(api_key=GEMINI_API_KEY)

# לוגים
extraction_log = []  # בעיות בחילוץ כללי
uncertain_log = []  # שורות לא בטוחות


def log_extraction_issue(filename, page, message, raw_text=""):
    """רישום בעיות בחילוץ כללי"""
    entry = {
        "שם קובץ": filename,
        "עמוד": page,
        "הודעה": message,
        "טקסט גולמי": raw_text
    }
    extraction_log.append(entry)
    print(f"[בעיה בחילוץ] {filename} (עמוד {page}): {message}")


def log_uncertain_row(filename, page, row_data):
    """רישום שורות לא בטוחות"""
    entry = {
        "שם קובץ": filename,
        "עמוד": page,
        "תאריך": row_data.get('date', ''),
        "שעה": row_data.get('time', ''),
        "מוצא": row_data.get('origin', ''),
        "תחנות": row_data.get('waypoints', ''),
        "יעד": row_data.get('destination', ''),
        "כמות": row_data.get('quantity', ''),
        "מחיר": f"{row_data.get('total_price_incl', '')} / {row_data.get('price_per_trip_incl', '')}",
        "הודעה": "המערכת לא בטוחה בנתונים שחולצו - בדוק ידנית",
        "טקסט גולמי": row_data.get('raw_text', '')
    }
    uncertain_log.append(entry)
    print(f"[שורה לא בטוחה] {filename} (עמוד {page})")


def make_excel_rtl(filepath):
    try:
        wb = load_workbook(filepath)
        for sheet in wb.worksheets:
            sheet.sheet_view.rightToLeft = True
        wb.save(filepath)
    except Exception as e:
        print(f"אזהרה: לא ניתן להפוך לRTL: {str(e)}")


def find_working_model():
    """מחפש מודל זמין - 2.5-flash קודם"""
    candidates = [
        "gemini-2.5-flash",  # החדש ביותר - ננסה ראשון
        "gemini-2.0-flash",
        "gemini-1.5-pro",
        "gemini-1.5-flash"
    ]
    print("מחפש מודל זמין...")
    for model in candidates:
        try:
            client.models.generate_content(model=model, contents="Test")
            print(f"✓ נבחר מודל: {model}")
            return model
        except Exception as e:
            print(f"  {model} - לא זמין")
            continue
    return None


def analyze_file_with_gemini(file_path, mime_type, active_model):
    """שולח קובץ לניתוח"""

    with open(file_path, 'rb') as f:
        file_bytes = f.read()

    prompt_text = """
אתה עוזר AI שחולץ נתונים מחשבוניות הסעות. קרא את הטבלה ותחלץ נתונים לפי התהליך הבא:

════════════════════════════════════════════════
תהליך חילוץ - עקוב בדיוק:
════════════════════════════════════════════════

לכל שורה בטבלה, בצע בסדר זה:

1️⃣ קרא את השורה הנוכחית
   - אל תשתמש במידע משורות קודמות
   - הסתכל רק על השורה הנוכחית

2️⃣ בדוק: זו כותרת או שורת נתונים?
   כותרת = יש בה רק PO/תאריך כללי ללא מסלול ספציפי
   אם כותרת → דלג על השורה

3️⃣ חלץ תאריך ושעה מהשורה הנוכחית
   פורמט: DD/MM/YYYY, HH:MM
   אם אין → null (זה בסדר)

4️⃣ חלץ מסלול מהשורה הנוכחית:

   A. מצא את העמודה עם תיאור המסלול
   B. קרא את הטקסט המדויק שכתוב שם
   C. פרק לפי הכלל:
      - מקום ראשון = origin
      - מקום אחרון = destination  
      - כל המקומות באמצע = waypoints (מופרד ב-;)

   **🔴 דוגמה קריטית - למד ממנה!**
   טבלה עם 17 שורות: "שריגים-כפר זוהרים" (חוזר 17 פעמים)
   שורה 18: "שריגים-**נבטים**" (שונה!)
   שורה 19: "נבטים-שריגים" (שונה!)
   שורה 20: "נבטים-גברעם-אמציה+שריגים" (שונה ומורכב!)

   **אתה חייב לקרוא כל שורה בדיוק כמו שהיא כתובה!**
   אם שורה 18 כתוב "שריגים-נבטים" - אל תכתוב "שריגים-כפר זוהרים" רק בגלל שזה היה ב-17 השורות הקודמות!

   סימני הפרדה בין מקומות: "-", "+", "/", "ל-"

   דוגמאות נוספות:

   טקסט: "שריגים-כפר זוהרים"
   → origin: "שריגים", waypoints: "", destination: "כפר זוהרים"

   טקסט: "שריגים-נבטים"  
   → origin: "שריגים", waypoints: "", destination: "נבטים"

   טקסט: "נבטים-גברעם-אמציה+שריגים"
   → origin: "נבטים", waypoints: "גברעם; אמציה", destination: "שריגים"

   טקסט: "קלחים-אמציה+שריגים"
   → origin: "קלחים", waypoints: "אמציה", destination: "שריגים"

   טקסט: "מרכזית עפולה הישנה-קיבוץ לביא צ.מגדל+חוקוק"
   → origin: "מרכזית עפולה הישנה", waypoints: "קיבוץ לביא; צ.מגדל", destination: "חוקוק"

   טקסט: "מיניבוס איסוף הבנות מנוב עצירת אסיפה מאבני איתן, והורדה במעיין עין עלמין"
   → origin: "נוב", waypoints: "אבני איתן", destination: "מעיין עין עלמין"

   ⚠️ חשוב מאוד:
   - **אל תמחק את המקום הראשון!** גם אם לפניו יש "מ-", "איסוף", "הבנות מ-"
   - מילים כמו "איסוף", "עצירת", "אסיפה" רק עוזרות לזהות את המקומות - אל תמחק מקומות!
   - "מנוב" → המקום הוא "נוב" (הסר "מ-")
   - "מאבני איתן" → המקום הוא "אבני איתן" (הסר "מ-")
   - "לחמרה" → המקום הוא "חמרה" (הסר "ל-")

   ניקוי שמות:
   - הסר רק את הקידומות: "מ-", "ל-", "ב-" מתחילת שמות
   - אל תמחק את המקום עצמו!

5️⃣ בדוק: האם זו תוספת מחיר?
   אם הטקסט מתחיל ב"תוספת" ויש תיאור (לא רק PO):
   → origin=null, waypoints="", destination=null

   אחרת: השתמש במסלול שחילצת בשלב 4

6️⃣ חלץ כמות מהשורה הנוכחית:
   חפש עמודת "כמות" או מספר לפני "נסיעות"
   אם אין → 1

7️⃣ חלץ מחירים מהשורה הנוכחית:
   ⚠️ **כלל זהב: אל תחשב כלום! רק העתק מספרים מעמודות!**

   - אם יש עמודה "מחיר יחידה לפני מע"מ" → price_per_trip_excl
   - אם יש עמודה "מחיר יחידה כולל מע"מ" → price_per_trip_incl
   - אם יש עמודה "סה"כ לפני מע"מ" → total_price_excl
   - אם יש עמודה "סה"כ כולל מע"מ" → total_price_incl

   **אם עמודה לא קיימת → null**
   **אל תחשב מע"מ! אם אין מחיר כולל מע"מ בעמודה → null**

8️⃣ בדוק פיצול:
   האם בשורה זו יש כמה תאריכים או מסלולים?
   אם כן → צור שורה לכל קומבינציה
   אם לא → שורה אחת בלבד

   ⚠️ חשוב: אם יש תאריך+שעה אחת בלבד = שורה אחת!
   דוגמה: "30/09/2025 13:15 קלחים-אמציה+שריגים" = שורה אחת

   💰 **מחירים בפיצול:**
   אם פיצלת שורה אחת למספר שורות:
   - כל השורות (חוץ מהאחרונה): price_per_trip_excl=null, price_per_trip_incl=null, total_price_excl=null, total_price_incl=null
   - רק השורה **האחרונה** תקבל את המחיר המלא
   - סמן split_from_combined=true בכל השורות

9️⃣ שמור טקסט גולמי מהשורה

════════════════════════════════════════════════
טעויות נפוצות - הימנע:
════════════════════════════════════════════════

❌ לא נכון: להשתמש במסלול משורה קודמת
✅ נכון: לקרוא את המסלול מהשורה הנוכחית

❌ לא נכון: לפצל שורה עם תאריך+שעה אחת
✅ נכון: שורה אחת = תאריך+שעה אחת

❌ לא נכון: לחשב מע"מ (524.58 × 1.18 = 619)
✅ נכון: רק להעתיק מה שכתוב בעמודה

❌ לא נכון: "מיניבוס איסוף מנוב" → למחוק את "נוב"
✅ נכון: "מיניבוס איסוף מנוב" → origin: "נוב"

❌ לא נכון: "קיבוץ לביא צ.מגדל" (יחד)
✅ נכון: "קיבוץ לביא; צ.מגדל" (מופרדות)

════════════════════════════════════════════════
פורמט JSON:
════════════════════════════════════════════════

{
    "rows": [
        {
            "supplier": "שם מראש העמוד",
            "page_number": מספר עמוד,
            "date": "DD/MM/YYYY" או null,
            "time": "HH:MM" או null,
            "origin": "שם נקי" או null,
            "waypoints": "שם1; שם2" או "",
            "destination": "שם נקי" או null,
            "quantity": מספר,
            "price_per_trip_excl": מספר או null,
            "price_per_trip_incl": מספר או null,
            "total_price_excl": מספר או null,
            "total_price_incl": מספר או null,
            "vat_found": true/false,
            "is_uncertain": true/false,
            "split_from_combined": true/false,
            "raw_text": "טקסט מלא מהשורה"
        }
    ]
}

════════════════════════════════════════════════
התחל עכשיו - חלץ את כל השורות
════════════════════════════════════════════════
    """

    for attempt in range(3):
        try:
            response = client.models.generate_content(
                model=active_model,
                contents=[types.Content(parts=[
                    types.Part.from_text(text=prompt_text),
                    types.Part.from_bytes(data=file_bytes, mime_type=mime_type)
                ])],
                config=types.GenerateContentConfig(response_mime_type='application/json')
            )

            data = json.loads(response.text)

            if isinstance(data, list):
                return {"rows": data}

            return data

        except Exception as e:
            if "429" in str(e):
                wait_time = 3 * (attempt + 1)
                print(f"Rate limit - ממתין {wait_time} שניות (ניסיון {attempt + 1}/3)")
                time.sleep(wait_time)
                continue
            else:
                if attempt < 2:
                    time.sleep(2)
                    continue
                return None

    return None


def main():
    print("=== תהליך 1: חילוץ נתונים מחשבוניות ===\n")

    if not os.path.exists(OUTPUT_FOLDER):
        os.makedirs(OUTPUT_FOLDER)
    if not os.path.exists(INPUT_FOLDER):
        print("שגיאה: תיקיית חשבוניות לא נמצאה.")
        return

    active_model = find_working_model()
    if not active_model:
        print("שגיאה: אין מודל זמין.")
        return

    all_data = []
    files = [f for f in os.listdir(INPUT_FOLDER) if f.lower().endswith(('.pdf', '.jpg', '.png', '.jpeg'))]
    print(f"נמצאו {len(files)} קבצים.\n")

    for i, filename in enumerate(files, 1):
        full_path = os.path.join(INPUT_FOLDER, filename)
        print(f"[{i}/{len(files)}] {filename}")

        mime_type = "application/pdf" if filename.lower().endswith(".pdf") else "image/jpeg"
        data = analyze_file_with_gemini(full_path, mime_type, active_model)

        if not data:
            log_extraction_issue(filename, "?", "Gemini לא החזיר תשובה (None) - אולי safety blocking או שגיאת API", "")
            continue

        if not data.get('rows'):
            response_str = json.dumps(data, ensure_ascii=False)[:500]
            log_extraction_issue(filename, "?", f"Gemini החזיר תשובה ריקה או ללא 'rows'. תשובה: {response_str}", "")
            continue

        rows = data.get('rows', [])
        print(f"  נמצאו {len(rows)} שורות")

        # קיבוץ שורות לפי ספק
        rows_by_supplier = {}
        for row in rows:
            supplier = row.get('supplier', 'לא זוהה')
            if supplier not in rows_by_supplier:
                rows_by_supplier[supplier] = []
            rows_by_supplier[supplier].append(row)

        # עיבוד כל ספק בנפרד
        for supplier, supplier_rows in rows_by_supplier.items():
            # חישוב סכומים לספק
            total_excl_vat = 0
            total_incl_vat = 0

            for idx, row in enumerate(supplier_rows):
                page_number = row.get('page_number', '?')
                date_val = row.get('date')
                time_val = row.get('time')
                origin = row.get('origin')
                waypoints = row.get('waypoints', '')
                dest = row.get('destination')
                quantity = row.get('quantity', 1)

                p_per_ex = row.get('price_per_trip_excl')
                p_per_in = row.get('price_per_trip_incl')
                p_tot_ex = row.get('total_price_excl')
                p_tot_in = row.get('total_price_incl')

                vat_found = row.get('vat_found', False)
                is_uncertain = row.get('is_uncertain', False)
                split_from_combined = row.get('split_from_combined', False)
                raw_text = row.get('raw_text', '')

                if is_uncertain:
                    log_uncertain_row(filename, page_number, row)

                rel_path = os.path.join("..", "חשבוניות", filename)
                excel_link = f'=HYPERLINK("{rel_path}", "פתח")'

                notes = []

                # הערות לשורות מפוצלות
                if split_from_combined:
                    if not p_tot_ex and not p_tot_in and not p_per_ex and not p_per_in:
                        notes.append("המחיר בשורה האחרונה")

                if not vat_found:
                    notes.append("מע״מ לא נמצא בחשבונית")
                if is_uncertain:
                    notes.append("ביטחון נמוך - בדוק ידנית")
                if not origin and not dest:
                    notes.append("תוספת מחיר")

                # צבירה לסכום כולל של הספק
                if p_tot_ex and p_tot_ex > 0:
                    total_excl_vat += p_tot_ex
                if p_tot_in and p_tot_in > 0:
                    total_incl_vat += p_tot_in

                # סכום כולל לספק - רק בשורה האחרונה
                invoice_total_excl = ""
                invoice_total_incl = ""
                if idx == len(supplier_rows) - 1:  # שורה אחרונה
                    # תמיד הצג סיכום, גם אם 0
                    invoice_total_excl = total_excl_vat

                    # אם יש סכום לפני מע"מ אבל אין כולל מע"מ - חשב
                    if total_excl_vat > 0 and total_incl_vat == 0:
                        invoice_total_incl = round(total_excl_vat * 1.18, 2)
                    else:
                        invoice_total_incl = total_incl_vat

                all_data.append({
                    "קובץ": filename,
                    "עמוד": page_number,
                    "קישור": excel_link,
                    "ספק": supplier,
                    "תאריך": date_val,
                    "שעה": time_val,
                    "מוצא": origin,
                    "תחנות": waypoints,
                    "יעד": dest,
                    "כמות": quantity,
                    "מחיר בודד לפני מע״מ": p_per_ex,
                    "מחיר בודד כולל מע״מ": p_per_in,
                    "מחיר סופי לפני מע״מ": p_tot_ex,
                    "מחיר סופי כולל מע״מ": p_tot_in,
                    "סה״כ חשבונית לפני מע״מ": invoice_total_excl,
                    "סה״כ חשבונית כולל מע״מ": invoice_total_incl,
                    "מע״מ נמצא": "כן" if vat_found else "לא",
                    "פוצל משורה מורכבת": "כן" if split_from_combined else "לא",
                    "הערות": ' | '.join(notes) if notes else "",
                    "טקסט גולמי": raw_text
                })

        time.sleep(2)

    timestamp = datetime.now().strftime('%Y%m%d_%H%M')

    # שמירת נתונים
    if all_data:
        df = pd.DataFrame(all_data)
        extracted_file = os.path.join(OUTPUT_FOLDER, f"נתונים_גולמיים_{timestamp}.xlsx")
        df.to_excel(extracted_file, index=False)
        make_excel_rtl(extracted_file)
        print(f"\n✓ נתונים גולמיים נשמרו: {extracted_file}")
        print(f"  סה״כ {len(all_data)} שורות חולצו")

    # לוגים
    log_file = os.path.join(OUTPUT_FOLDER, f"לוג_חילוץ_{timestamp}.xlsx")

    with pd.ExcelWriter(log_file, engine='openpyxl') as writer:
        if extraction_log:
            df_extraction = pd.DataFrame(extraction_log)
            df_extraction.to_excel(writer, sheet_name='בעיות חילוץ', index=False)
        else:
            df_empty = pd.DataFrame({"הודעה": ["לא נמצאו בעיות בחילוץ"]})
            df_empty.to_excel(writer, sheet_name='בעיות חילוץ', index=False)

        if uncertain_log:
            df_uncertain = pd.DataFrame(uncertain_log)
            df_uncertain.to_excel(writer, sheet_name='שורות לא בטוחות', index=False)
        else:
            df_empty = pd.DataFrame({"הודעה": ["לא נמצאו שורות לא בטוחות"]})
            df_empty.to_excel(writer, sheet_name='שורות לא בטוחות', index=False)

    make_excel_rtl(log_file)
    print(f"✓ לוג חילוץ נשמר: {log_file}")
    print(f"  - בעיות חילוץ: {len(extraction_log)}")
    print(f"  - שורות לא בטוחות: {len(uncertain_log)}")

    print("\n=== תהליך חילוץ הסתיים בהצלחה! ===")
    print("קובץ הנתונים הגולמיים מוכן לשלב 2 (עיבוד וחישובים)")


if __name__ == "__main__":
    main()