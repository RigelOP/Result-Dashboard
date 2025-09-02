#!/usr/bin/env python3


import os
import time
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
from bs4 import BeautifulSoup
import re
from datetime import datetime
import logging

# module logger (do not configure handlers here so importing apps control logging)
logger = logging.getLogger(__name__)

# Configuration
EXCEL_FILE_PATH = "student_data.xlsx"
RESULT_URL = "https://result.mdurtk.in/postexam/result.aspx"

# Browser window configuration
FIXED_WINDOW_SIZE = False  # Set to False to use maximize window
WINDOW_WIDTH = 1366       # Fixed window width (when FIXED_WINDOW_SIZE is True)
WINDOW_HEIGHT = 768       # Fixed window height (when FIXED_WINDOW_SIZE is True)
ZOOM_LEVEL = '67%'        # Zoom level to maintain throughout the program

def wait_for_element(driver, by, value, timeout=10):
    """Wait for an element to be present"""
    try:
        return WebDriverWait(driver, timeout).until(
            EC.presence_of_element_located((by, value))
        )
    except:
        return None

def wait_for_page_load(driver, timeout=10):
    """Wait for page to be fully loaded"""
    try:
        WebDriverWait(driver, timeout).until(
            lambda d: d.execute_script("return document.readyState") == "complete"
        )
        return True
    except:
        return False

def wait_for_result_content(driver, timeout=15):
    """Wait for result content to load properly"""
    try:
        # Wait for result tables to be present
        WebDriverWait(driver, timeout).until(
            lambda d: len(d.find_elements(By.TAG_NAME, "table")) > 0
        )
        # Additional wait for content to stabilize
        time.sleep(3)
        return True
    except:
        return False



def extract_data_from_page(driver):
    """Extract data from result page including numerical marks, F grades, and number+F combinations (like 29F)"""
    try:
        page_source = driver.page_source
        soup = BeautifulSoup(page_source, 'html.parser')
        
        # Find tables with marks
        marks_data = {}
        tables = soup.find_all('table')
        
        for table in tables:
            table_text = table.get_text().upper()
            if 'TOTAL' in table_text or 'MARKS' in table_text:
                rows = table.find_all('tr')
                
                for row in rows[1:]:  # Skip header
                    cells = row.find_all(['td', 'th'])
                    if len(cells) >= 3:
                        first_cell = cells[0].get_text().strip()
                        
                        # Look for subject patterns
                        subject_patterns = [
                            r'^([A-Z]{2,3}-[A-Z]{2,4}-[0-9]{2,3}[A-Z]?)[\s:]',
                            r'^([A-Z]{2,3}-[A-Z]{2,4}-[IVX]+)[\s:]',
                            r'^([A-Z]{2,3}-[0-9]{2,3}[A-Z]?)[\s:]'
                        ]
                        
                        for pattern in subject_patterns:
                            match = re.search(pattern, first_cell)
                            if match:
                                subject_code = match.group(1)
                                
                                # Extract marks (including F grades and number+F combinations)
                                row_text = ' '.join([cell.get_text().strip() for cell in cells])
                                
                                # First look for number+F pattern (like 29F, 15F, etc.)
                                number_f_pattern = r'\b([0-9]{1,3}F)\b'
                                number_f_match = re.search(number_f_pattern, row_text, re.IGNORECASE)
                                if number_f_match:
                                    marks_data[subject_code] = number_f_match.group(1).upper()
                                else:
                                    # Look for standalone F grade
                                    f_pattern = r'\bF\b'
                                    if re.search(f_pattern, row_text, re.IGNORECASE):
                                        marks_data[subject_code] = 'F'
                                    else:
                                        # Look for numerical marks
                                        numbers = re.findall(r'\b([0-9]{2,3})\b', row_text)
                                        valid_marks = [n for n in numbers if 10 <= int(n) <= 100]
                                        
                                        if valid_marks:
                                            marks_data[subject_code] = valid_marks[-1]
                                        else:
                                            # Check for other fail indicators
                                            fail_patterns = [r'\bFAIL\b', r'\bAB\b', r'\bABSENT\b']
                                            for fail_pattern in fail_patterns:
                                                if re.search(fail_pattern, row_text, re.IGNORECASE):
                                                    marks_data[subject_code] = 'F'
                                                    break
                                break
        
        return marks_data
    except Exception:
        logger.exception("Error extracting data from page")
        return {}

def process_single_student(driver, name, roll, reg):
    """Process one student"""
    safe_name = name.replace(" ", "_").replace("/", "_")
    logger.info("Processing: %s (Roll: %s)", name, roll)
    
    try:
        # Fill form
        logger.debug("Filling form for %s", name)
        roll_input = wait_for_element(driver, By.ID, "txtRollNo")
        if not roll_input:
            raise Exception("Roll input not found")

        roll_input.clear()
        roll_input.send_keys(roll)

        reg_input = wait_for_element(driver, By.ID, "txtRegistrationNo")
        if not reg_input:
            raise Exception("Registration input not found")

        reg_input.clear()
        reg_input.send_keys(reg)

        # Submit
        submit_btn = wait_for_element(driver, By.ID, "cmdbtnProceed")
        if not submit_btn:
            raise Exception("Submit button not found")

        submit_btn.click()
        wait_for_page_load(driver)

        # Confirm
        logger.debug("Clicking confirm if present")
        for confirm_name in ["imgComfirm", "cmdconfirm"]:
            confirm_btn = wait_for_element(driver, By.NAME, confirm_name, 5)
            if confirm_btn:
                confirm_btn.click()
                break

        wait_for_page_load(driver)

        # View result
        logger.debug("Opening result view")
        view_link = wait_for_element(driver, By.LINK_TEXT, "View")
        if not view_link:
            raise Exception("View link not found")

        view_link.click()
        wait_for_page_load(driver)

        # Wait for result content to load completely
        logger.debug("Waiting for result content to load")
        if not wait_for_result_content(driver):
            logger.warning("Result content may not be fully loaded")

        # Extract data
        logger.debug("Extracting data from result page")
        marks_data = extract_data_from_page(driver)
        logger.debug("Found %d subjects", len(marks_data))

        logger.info("SUCCESS - %s", name)
        return {
            'name': name,
            'roll': roll,
            'reg': reg,
            'status': 'Success',
            'marks': marks_data
        }

    except Exception as e:
        logger.exception("Error processing student %s (%s)", name, roll)
        return {
            'name': name,
            'roll': roll,
            'reg': reg,
            'status': f'Failed: {e}',
            'marks': {}
        }

def collect_subject_names(all_subjects):
    """Collect subject names and total marks from user input"""
    logger.info("Starting subject information collection for %d subjects", len(all_subjects))

    subject_info = {}
    for i, subject_code in enumerate(sorted(all_subjects), 1):
        logger.info("%d. Subject Code: %s", i, subject_code)
        
        # Get subject name
        while True:
            subject_name = input(f"   Enter full name for {subject_code}: ").strip()
            if subject_name:
                break
            logger.warning("Subject name cannot be empty for %s", subject_code)
        
        # Get total marks
        while True:
            try:
                total_marks = input(f"   Enter total marks for {subject_code}: ").strip()
                total_marks = int(total_marks)
                if total_marks > 0:
                    break
                else:
                    logger.warning("Total marks must be a positive number for %s", subject_code)
            except ValueError:
                logger.warning("Invalid number entered for total marks for %s", subject_code)
        
        subject_info[subject_code] = {
            'name': subject_name,
            'total_marks': total_marks
        }
    
    return subject_info

def create_grade_distribution_analysis(results, subject_info):
    """Create grade distribution analysis for each subject"""
    if not results or not subject_info:
        return []
    
    # Define grade ranges
    grade_ranges = [
        ('Above 80%', 80, 100),
        ('70-80%', 70, 79),
        ('60-70%', 60, 69),
        ('50-60%', 50, 59),
        ('Below 50%', 0, 49)
    ]
    
    analysis_data = []
    
    for subject_code, info in subject_info.items():
        subject_name = info['name']
        total_marks = info['total_marks']
        
        # Count students in each grade range
        range_counts = {}
        for range_name, min_pct, max_pct in grade_ranges:
            range_counts[range_name] = 0
        
        # Analyze each student's marks for this subject
        for result in results:
            if result['marks'] and subject_code in result['marks']:
                mark_str = str(result['marks'][subject_code])
                
                # Skip if it contains 'F' (failed grades)
                if 'F' in mark_str.upper() or mark_str == 'N/A':
                    continue
                
                try:
                    mark = int(mark_str)
                    percentage = (mark / total_marks) * 100
                    
                    # Categorize into ranges (using lower limit logic)
                    if percentage >= 80:
                        range_counts['Above 80%'] += 1
                    elif percentage >= 70:
                        range_counts['70-80%'] += 1
                    elif percentage >= 60:
                        range_counts['60-70%'] += 1
                    elif percentage >= 50:
                        range_counts['50-60%'] += 1
                    else:
                        range_counts['Below 50%'] += 1
                        
                except ValueError:
                    continue
        
        # Create row for this subject
        row = {
            'Subject_Code': subject_code,
            'Subject_Name': subject_name,
            'Total_Marks': total_marks
        }
        
        # Add counts for each range
        for range_name, _, _ in grade_ranges:
            row[range_name] = range_counts[range_name]
        
        analysis_data.append(row)
    
    return analysis_data

def save_results_to_excel(results, filename, subject_info=None):
    """Save results to Excel with optional grade distribution analysis"""
    if not results:
        logger.warning("No results to save to Excel: %s", filename)
        return
    
    # Get all subjects
    all_subjects = set()
    for result in results:
        if result['marks']:
            all_subjects.update(result['marks'].keys())
    
    all_subjects = sorted(list(all_subjects))
    
    # Create Excel data
    excel_data = []
    for result in results:
        row = {
            'Student_Name': result['name'],
            'Roll_Number': result['roll'],
            'Registration_Number': result['reg'],
            'Status': result['status']
        }
        
        # Add marks for each subject
        for subject in all_subjects:
            row[subject] = result['marks'].get(subject, 'N/A')
        
        excel_data.append(row)
    
    # Save to Excel with multiple sheets
    with pd.ExcelWriter(filename, engine='openpyxl') as writer:
        # Main results sheet
        df = pd.DataFrame(excel_data)
        df.to_excel(writer, sheet_name='Student_Results', index=False)
        
        # Grade distribution analysis sheet (if subject info provided)
        if subject_info:
            analysis_data = create_grade_distribution_analysis(results, subject_info)
            if analysis_data:
                analysis_df = pd.DataFrame(analysis_data)
                analysis_df.to_excel(writer, sheet_name='Grade_Distribution', index=False)
                logger.info("Grade distribution analysis added to Excel: %s", filename)
    
    logger.info("Results saved to: %s", filename)

# Main execution
def run_cli(excel_path=EXCEL_FILE_PATH, interactive=True):
    """Run the full scraping CLI flow. Parameters kept for convenient invocation from other scripts/tests."""
    try:
        # Load student data
        logger.info("Loading student data from %s", excel_path)
        df = pd.read_excel(excel_path)
        logger.info("Loaded %d students", len(df))

        # Build student list
        students = []
        for i in range(len(df)):
            row = df.iloc[i]
            if not (pd.isna(row['Roll Number']) or pd.isna(row['Registration Number']) or pd.isna(row['Student Name'])):
                students.append({
                    'name': str(row['Student Name']).strip(),
                    'roll': str(int(row['Roll Number'])),
                    'reg': str(int(row['Registration Number']))
                })

        logger.info("Processing ALL %d students from Excel file", len(students))

        # Setup browser
        logger.info("Setting up browser for scraping")
        options = Options()
        options.add_argument("--disable-gpu")
        options.add_argument("--no-sandbox")
        options.add_argument("--disable-infobars")
        options.add_argument("--disable-extensions")
        options.add_experimental_option("excludeSwitches", ["enable-automation"])
        options.add_experimental_option('useAutomationExtension', False)

        driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)

        try:
            # Open website
            logger.info("Opening %s", RESULT_URL)
            driver.get(RESULT_URL)

            # Configure window size
            if FIXED_WINDOW_SIZE:
                logger.info("Setting fixed window size: %dx%d", WINDOW_WIDTH, WINDOW_HEIGHT)
                driver.set_window_size(WINDOW_WIDTH, WINDOW_HEIGHT)
            else:
                logger.info("Maximizing browser window")
                driver.maximize_window()

            # Set zoom level for the entire session
            logger.info("Setting zoom level to %s", ZOOM_LEVEL)
            driver.execute_script(f"document.body.style.zoom='{ZOOM_LEVEL}'")

            # Remove automation detection
            driver.execute_script("Object.defineProperty(navigator, 'webdriver', {get: () => undefined})")

            wait_for_page_load(driver)

            # Process students
            all_results = []
            successful = 0
            failed = 0

            for i, student in enumerate(students, 1):
                logger.info("Student %d/%d: %s", i, len(students), student['name'])
                result = process_single_student(driver, student['name'], student['roll'], student['reg'])
                all_results.append(result)

                if result['status'] == 'Success':
                    successful += 1
                    logger.info("SUCCESS - total successful: %d", successful)
                else:
                    failed += 1
                    logger.info("FAILED - total failed: %d", failed)

                logger.debug("Progress: %d successful, %d failed", successful, failed)

                # Navigate back for next student
                if i < len(students):
                    logger.debug("Re-loading result page for next student")
                    driver.get(RESULT_URL)
                    wait_for_page_load(driver)
                    driver.execute_script(f"document.body.style.zoom='{ZOOM_LEVEL}'")
                    time.sleep(1)

            # Save results
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            filename = f"Complete_Results_{timestamp}.xlsx"

            # Get all subjects found
            all_subjects = set()
            for result in all_results:
                if result['marks']:
                    all_subjects.update(result['marks'].keys())

            # Collect subject information from user if interactive
            subject_info = None
            if all_subjects and interactive:
                logger.info("Found %d unique subjects", len(all_subjects))
                collect_info = input("Would you like to provide subject names and total marks? (y/n): ").strip().lower()
                if collect_info in ['y', 'yes']:
                    subject_info = collect_subject_names(all_subjects)
                else:
                    logger.info("Skipping subject information collection")

            save_results_to_excel(all_results, filename, subject_info)

            # Stats
            logger.info("PROCESSING COMPLETED")
            logger.info("Total students in Excel: %d", len(df))
            logger.info("Students processed: %d", len(students))
            logger.info("Successful extractions: %d", successful)
            logger.info("Failed extractions: %d", failed)
            logger.info("Complete results saved to: %s", filename)

        finally:
            driver.quit()
            logger.info("Browser closed")

    except Exception:
        logger.exception("Error in run_cli")


if __name__ == "__main__":
    # Run CLI by default when executed directly
    run_cli()
