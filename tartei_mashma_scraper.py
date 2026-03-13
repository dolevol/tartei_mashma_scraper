import undetected_chromedriver as uc
from bs4 import BeautifulSoup
import pandas as pd
from datetime import datetime, timedelta
import time
import re
import os
import time
import undetected_chromedriver as uc

def setup_stealth_driver():
    print("Cleaning up leftover browser processes...")
    # סגירה אוטומטית של כל תהליך כרום/דרייבר שנותר פתוח
    os.system("pkill -f chrome")
    os.system("pkill -f chromedriver")
    time.sleep(2) # המתנה קצרה לווידוא שהתהליכים נסגרו
    
    options = uc.ChromeOptions()
    # הכוונה מדויקת לקובץ כדי למנוע שגיאות ב-Codespaces
    options.binary_location = "/usr/bin/google-chrome" 
    
    options.add_argument('--no-sandbox')
    options.add_argument('--disable-dev-shm-usage')
    
    # אתחול הדרייבר: מצב headless מוגדר כאן
    driver = uc.Chrome(options=options, headless=True)
    return driver

def scrape_crosswords_hybrid():
    start_date = datetime.strptime("08/08/2025", "%d/%m/%Y")
    end_date = datetime.strptime("13/03/2026", "%d/%m/%Y")
    
    base_url = "https://www.14across.co.il/answers.php?wantsold=1&crossword=12&name=%D7%AA%D7%A8%D7%AA%D7%99%20%D7%9E%D7%A9%D7%9E%D7%A2,%20%D7%93%D7%A7%D7%9C%20%D7%91%D7%A0%D7%95&date={}"
    
    all_data = []
    current_date = start_date
    
    print("Initializing Stealth Browser...")
    driver = setup_stealth_driver()

    try:
        while current_date <= end_date:
            date_str = current_date.strftime("%d/%m/%Y")
            url = base_url.format(date_str)
            print(f"Fetching data for: {date_str}...")
            
            driver.get(url)
            
            max_wait = 15
            for _ in range(max_wait):
                if "sgcaptcha" not in driver.page_source:
                    break
                time.sleep(1)
            else:
                print(f"  -> Failed to bypass CAPTCHA for date {date_str}.")
                driver.save_screenshot(f"captcha_blocked_{date_str.replace('/','')}.png")
                current_date += timedelta(days=7)
                continue
            
            time.sleep(2)
            
            html_source = driver.page_source
            soup = BeautifulSoup(html_source, 'html.parser')
            
            title_tag = soup.find('h2', string=lambda text: text and 'כותרת התשבץ' in text)
            crossword_title = title_tag.text.replace("כותרת התשבץ:", "").strip() if title_tag else "Unknown"

            answers_content = soup.find('div', id='answers-content')
            
            if not answers_content:
                print("  -> No crossword content found on this page.")
                current_date += timedelta(days=7) 
                continue
                
            question_paragraphs = answers_content.find_all('p')
            
            for p in question_paragraphs:
                q_num_span = p.find('span', class_='question_number')
                if not q_num_span:
                    continue
                    
                def_number = q_num_span.text.replace(":", "").strip()
                ans_span = p.find('span', class_='actual-answer')
                solution = ans_span['data-content'].strip() if ans_span and 'data-content' in ans_span.attrs else ""
                
                # --- Updated Explanation Parsing Logic ---
                explanation = ""
                
                # Extract the unique ID from the link tag (which correctly stays inside the <p>)
                link_tag = p.find('a', class_='answer-links')
                if link_tag and 'id' in link_tag.attrs:
                    struct_id = link_tag['id'].replace('link-', '')
                    
                    # Search globally in the 'soup' using the exact ID, bypassing the <p> tag limitation
                    help_texts_div = soup.find('div', id=f'help-text-{struct_id}')
                    
                    if help_texts_div:
                        ul_tag = help_texts_div.find('ul')
                        if ul_tag:
                            first_li = ul_tag.find('li')
                            if first_li:
                                explanation = first_li.get_text(strip=True)
                # -----------------------------------------
                
                all_data.append({
                    "כותרת התשבץ": crossword_title,
                    "תאריך התשבץ": date_str,
                    "מספר הגדרה": def_number,
                    "פתרון": solution,
                    "הסבר": explanation
                })

            current_date += timedelta(days=7)
            
    finally:
        driver.quit()
        
    if all_data:
        df = pd.DataFrame(all_data)
        output_filename = f"crosswords_hybrid_{start_date.strftime('%Y%m%d')}_to_{end_date.strftime('%Y%m%d')}.xlsx"
        df.to_excel(output_filename, index=False, engine='openpyxl')
        print(f"\nScraping completed successfully! Saved {len(df)} records to: {output_filename}")
    else:
        print("\nNo data was collected.")

if __name__ == "__main__":
    scrape_crosswords_hybrid()