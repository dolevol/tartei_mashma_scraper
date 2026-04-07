import os
import re
from datetime import datetime
from pypdf import PdfWriter

def extract_date_from_filename(filename):
    """
    מחלץ תאריך משם הקובץ ומחזיר אובייקט datetime למיון.
    אם לא נמצא תאריך, מחזיר תאריך מקסימלי כדי שיופיע בסוף (למרות שהקבצים המיוחדים מטופלים בנפרד).
    """
    match = re.search(r'(\d{1,2})\.(\d{1,2})\.(\d{2,4})', filename)
    
    if match:
        day = int(match.group(1))
        month = int(match.group(2))
        year = int(match.group(3))
        
        if year < 100:
            year += 2000
            
        try:
            return datetime(year, month, day)
        except ValueError:
            pass 
            
    return datetime.max

def merge_pdfs_with_covers(folder_path, output_filename):
    """
    סורקת תיקייה, מוצאת קבצי PDF, ממיינת כרונולוגית,
    ומוסיפה קובץ פתיחה וקובץ סיום ספציפיים (אם הם קיימים בתיקייה).
    """
    # הגדרת השמות המיוחדים (בדיוק כפי שרשמת)
    START_FILE = "כריכה ותוכן עניינים.pdf"
    END_FILE = "crossword_solutions.pdf"
    
    # 1. איסוף כל קבצי ה-PDF בתיקייה
    all_pdf_files = [f for f in os.listdir(folder_path) if f.lower().endswith('.pdf')]
    
    if not all_pdf_files:
        print(f"לא נמצאו קבצי PDF בתיקייה: {folder_path}")
        return

    # 2. הפרדת הקבצים המיוחדים ושאר הקבצים למיון
    regular_files = []
    has_start_file = False
    has_end_file = False
    
    for f in all_pdf_files:
        if f == START_FILE:
            has_start_file = True
        elif f == END_FILE:
            has_end_file = True
        else:
            regular_files.append(f)

    # 3. מיון הקבצים הרגילים (אלו עם התאריכים)
    sorted_files = sorted(regular_files, key=extract_date_from_filename)
    
    merger = PdfWriter()
    
    print("** מתחיל במיזוג הקבצים לפי הסדר הבא:**")
    
    # --- הוספת קובץ פתיחה ---
    if has_start_file:
        print(f" מצרף קובץ פתיחה: {START_FILE}")
        filepath = os.path.join(folder_path, START_FILE)
        try:
            merger.append(filepath)
        except Exception as e:
            print(f"שגיאה בצירוף קובץ הפתיחה {START_FILE}: {e}")
    else:
        print(f" אזהרה: קובץ פתיחה '{START_FILE}' לא נמצא בתיקייה, ממשיך בלעדיו.")
        
    # --- הוספת קבצים ממוינים ---
    for filename in sorted_files:
        file_date = extract_date_from_filename(filename)
        date_display = file_date.strftime("%d/%m/%Y") if file_date != datetime.max else "ללא תאריך מזוהה"
        print(f" מצרף את: {filename} \t(תאריך: {date_display})")
        
        filepath = os.path.join(folder_path, filename)
        try:
            merger.append(filepath)
        except Exception as e:
            print(f"שגיאה בצירוף הקובץ {filename}: {e}")
            
    # --- הוספת קובץ סיום ---
    if has_end_file:
        print(f" מצרף קובץ סיום: {END_FILE}")
        filepath = os.path.join(folder_path, END_FILE)
        try:
            merger.append(filepath)
        except Exception as e:
            print(f"שגיאה בצירוף קובץ הסיום {END_FILE}: {e}")
    else:
        print(f" אזהרה: קובץ סיום '{END_FILE}' לא נמצא בתיקייה, ממשיך בלעדיו.")

    # 4. שמירת הקובץ המאוחד - מחוץ לתיקייה, באותה רמה של סקריפט ה-Python
    # ה-os.path.dirname(__file__) ימצא את המיקום של קובץ הפייתון הזה
    current_dir = os.path.dirname(os.path.abspath(__file__))
    output_path = os.path.join(current_dir, output_filename)
    
    try:
        merger.write(output_path)
        print(f"\nהמיזוג הושלם בהצלחה! הקובץ נשמר מחוץ לתיקייה בכתובת:")
        print(f"--> {output_path}")
    except Exception as e:
         print(f"\nשגיאה בשמירת הקובץ המאוחד: {e}")
    finally:
        merger.close()

# --- אזור הרצה ---
if __name__ == "__main__":
    # הנתיב שבו נמצאים כל קבצי ה-PDF (כולל קבצי התאריכים וכולל כריכה/פתרונות)
    FOLDER_PATH = r"crossword_pdfs" 
    
    # השם של קובץ ה-PDF הסופי שייווצר
    OUTPUT_FILE = "Final_Merged_Crosswords.pdf"
    
    merge_pdfs_with_covers(FOLDER_PATH, OUTPUT_FILE)