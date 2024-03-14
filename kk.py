from fuzzywuzzy import fuzz
import requests
from bs4 import BeautifulSoup
import re
import speech_recognition as sr
import pyttsx3
import win32com.client as win32
from docx import Document
import os

# Initialize the recognizer
recognizer = sr.Recognizer()

# Initialize the text-to-speech engine
engine = pyttsx3.init()

def speak(text):
    engine.say(text)
    engine.runAndWait()

def is_similar(keyword, text): 
    return fuzz.ratio(keyword, text) >= 80  # Adjust the similarity threshold as needed

def clean_course_name(course_name):
    # Use regular expressions to clean the course name
    match = re.search(r'\((.*?)\)', course_name)  # Extract text within round brackets
    if match:
        cleaned_name = match.group(1)  # Extract the text within round brackets
    else:
        cleaned_name = course_name  # Use the original course name if no round brackets found

    cleaned_name = cleaned_name.strip()  # Remove extra spaces
    cleaned_name = cleaned_name.lower()  # Convert to lowercase for easier matching
    return cleaned_name

def scrape_website(url, course_class, fee_class, cutoff_class):
    try:
        # Set a User-Agent header
        headers = {
            "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/58.0.3029.110 Safari/537.36"
        }

        # Create a session to maintain headers and cookies
        session = requests.Session()
        response = session.get(url, headers=headers)

        # Check if the request was successful
        if response.status_code == 200:
            # Parse the HTML content of the page
            soup = BeautifulSoup(response.text, 'html.parser')

            # Find elements with the specified class names for course names, fees, and cut-offs
            course_elements = soup.find_all(class_=course_class)
            fee_elements = soup.find_all(class_=fee_class)
            cutoff_elements = soup.find_all(class_=cutoff_class)

            if course_elements and fee_elements and cutoff_elements:
                # Extract and return the course names, fees, and cut-offs as lists
                course_data = []
                cutoff_data = []  # Separate list for cutoff information
                for course, fee, cutoff in zip(course_elements, fee_elements, cutoff_elements):
                    course_name = clean_course_name(course.text)
                    fee_text = fee.text
                    cutoff_text = cutoff.text
                    course_data.append((course_name, fee_text, cutoff_text))
                    cutoff_data.append((course_name, cutoff_text))
                return course_data, cutoff_data
            else:
                return None
        else:
            return None

    except Exception as e:
        return None

def scrape_ranking(url, ranking_class):
    try:
        # Set a User-Agent header
        headers = {
            "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/58.0.3029.110 Safari/537.36"
        }

        # Create a session to maintain headers and cookies
        session = requests.Session()
        response = session.get(url, headers=headers)

        # Check if the request was successful
        if response.status_code == 200:
            # Parse the HTML content of the page
            soup = BeautifulSoup(response.text, 'html.parser')

            # Find elements with the specified class for ranking
            ranking_element = soup.find(class_=ranking_class)

            if ranking_element:
                # Extract and return the ranking text
                ranking_text = ranking_element.text.strip()
                return ranking_text
            else:
                return None
        else:
            return None

    except Exception as e:
        return None

def scrape_comparison(url, class_name, output_word_file, output_pdf_file):
    try:
        # Set a User-Agent header
        headers = {
            "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/58.0.3029.110 Safari/537.36"
        }

        # Create a session to maintain headers and cookies
        session = requests.Session()
        response = session.get(url, headers=headers)

        # Check if the request was successful
        if response.status_code == 200:
            # Parse the HTML content of the page
            soup = BeautifulSoup(response.text, 'html.parser')

            # Find the table element with the specified class name
            table = soup.find('table', class_=class_name)

            if table:
                # Create a Word document
                doc = Document()

                # Extract and add table data to the Word document
                for row in table.find_all('tr'):
                    row_data = [cell.get_text(strip=True) for cell in row.find_all(['th', 'td'])]
                    table = doc.add_table(rows=1, cols=len(row_data))
                    table.autofit = False
                    table.allow_autofit = False
                    table.style = 'Table Grid'
                    table.alignment = 1  # Left-align the table
                    for index, cell_text in enumerate(row_data):
                        cell = table.cell(0, index)
                        cell.text = cell_text

                # Save the Word document
                doc.save(output_word_file)

                # Convert Word to PDF
                word = win32.gencache.EnsureDispatch('Word.Application')
                doc = word.Documents.Open(output_word_file)
                output_pdf_file = output_pdf_file.replace('.docx', '.pdf')
                doc.SaveAs(output_pdf_file, FileFormat=17)
                doc.Close()
                word.Quit()

                print(f"Comparison data has been scraped and saved to '{output_pdf_file}' as a PDF.")
            else:
                print(f"No table found with class '{class_name}' on the page.")
        else:
            print(f"Failed to retrieve the page. Status code: {response.status_code}")

    except Exception as e:
        print(f"An error occurred: {str(e)}")

def main():
    print("Listening for your command...")
    
    downloads_dir = os.path.join(os.path.expanduser("~"), "Downloads")
    conversation_state = "initial"  # Initial state
    course_data = None  # Store course, fee, and cutoff information
    ranking_text = None  # Store ranking information
    comparison_data = None  # Store comparison table information
    jarvis_enabled = False  # Track if "jarvis" is mentioned
    response_text = "Not initiated"

    while True:
        with sr.Microphone() as source:
            recognizer.adjust_for_ambient_noise(source)  # Adjust for noise
            audio = recognizer.listen(source, timeout=5)  # Listen for 5 seconds

        try:
            # Recognize the speech
            command = recognizer.recognize_google(audio).lower()

            # Check if "jarvis" is mentioned
            if "jarvis" in command:
                response_text = "How can I assist you for the Admission Query?"
                course_class = "jsx-1946297154 text-title text-lg font-weight-bold mb-0 pr-5"
                fee_class = "jsx-1946297154 fee text-success font-weight-bold mr-1 text-lg"
                cutoff_class = "jsx-1946297154 course-details-item-data font-weight-medium"
                ranking_class = "jsx-2427240194 jsx-367356671"  # Class for ranking
                class_name = "jsx-3988704578 mb-0"
                url_course = "https://collegedunia.com/college/28749-pandit-deendayal-energy-university-school-of-technology-pdeu-sot-gandhinagar/courses-fees?slug=bachelor-of-technology-btech&course_type=Full-Time"
                url_ranking = "https://collegedunia.com/college/28749-pandit-deendayal-energy-university-school-of-technology-pdeu-sot-gandhinagar/courses-fees?course_id=2047"
                course_data, cutoff_data = scrape_website(url_course, course_class, fee_class, cutoff_class)
                ranking_text = scrape_ranking(url_ranking, ranking_class)
                jarvis_enabled = True
            elif jarvis_enabled:
                if course_data:
                    if "document" in command or "documents" in command and any("pdeu college" or "pdpu college" or course in command for course, _ in cutoff_data):
                        response_text = "The below-mentioned documents need to be attached to the online application:\n\n"
                        response_text += "1. HSC mark sheet\n"
                        response_text += "2. JEE Main admit card\n"
                        response_text += "3. JEE rank card\n"
                        response_text += "4. Caste certificate\n"
                        response_text += "5. Scanned passport-size photograph\n"

                    elif "eligibility" in command and any(course in command for course, _ in cutoff_data):
                        response_text = "Academic Requirement Candidate should have passed the Qualifying Examination with minimum 45% marks (40% in case of SC / ST) in aggregate in theory and practical of Physics & Maths (with Chemistry or Biology or Computer or Vocational Subject) from a single board."

                    elif "comparison for computer science" in command:
                    # Extract and save the comparison table
                            url_comparison_computer = "https://collegedunia.com/college-compare?entity_1=28749,2047&entity_2=25483,2047&popup=true"
                            output_comparison_word_file = os.path.join(downloads_dir, "comparison_data_computer.docx")
                            output_comparison_pdf_file = os.path.join(downloads_dir, "comparison_data_computer.pdf")
                            scrape_comparison(url_comparison_computer, class_name, output_comparison_word_file, output_comparison_pdf_file)
                            response_text = f"Comparison data has been downloaded for {command} as a Word and PDF file."

                    elif "comparison for ict" in command:
                        # Extract and save the comparison table
                        
                            url_comparison_ICT = "https://collegedunia.com/college-compare?entity_1=28749,5120&entity_2=25484,5120&popup=true"
                            output_comparison_word_file = os.path.join(downloads_dir, "comparison_data_ict.docx")
                            output_comparison_pdf_file = os.path.join(downloads_dir, "comparison_data_ict.pdf")
                            scrape_comparison(url_comparison_ICT, class_name, output_comparison_word_file, output_comparison_pdf_file)
                            response_text = f"Comparison data has been downloaded for {command} as a Word and PDF file."
                    
                    elif "comparison for mechanical engineering" in command:
                        # Extract and save the comparison table                        
                        url_comparison_mechanical = "https://collegedunia.com/college-compare?entity_1=28749,2094&entity_2=25997,2094&entity_3=25503,2094&entity_4=25914,2094&popup=true"
                        output_comparison_word_file = os.path.join(downloads_dir, "comparison_data_mechanical.docx")
                        output_comparison_pdf_file = os.path.join(downloads_dir, "comparison_data_mechanical.pdf")
                        scrape_comparison(url_comparison_mechanical, class_name, output_comparison_word_file,output_comparison_pdf_file)
                        response_text = f"Comparison data has been downloaded for {command} as a Word and PDF file."
                    
                    elif "comparison for chemical engineering" in command:
                        # Extract and save the comparison table                        
                        url_comparison_chemical = "https://collegedunia.com/college-compare?entity_1=28749,1937&entity_2=25997,1937&entity_3=25503,1937&entity_4=25914,1937&popup=true"
                        output_comparison_word_file = os.path.join(downloads_dir, "comparison_data_chemical.docx")
                        output_comparison_pdf_file = os.path.join(downloads_dir, "comparison_data_chemical.pdf")
                        scrape_comparison(url_comparison_chemical, class_name, output_comparison_word_file,output_comparison_pdf_file)
                        response_text = f"Comparison data has been downloaded for {command} as a Word and PDF file."

                    elif "comparison for civil engineering" in command:
                        # Extract and save the comparison table                        
                        url_comparison_civil = "https://collegedunia.com/college-compare?entity_1=28749,1938&entity_2=25997,1938&entity_3=25503,1938&entity_4=25914,1938&popup=true"
                        output_comparison_word_file = os.path.join(downloads_dir, "comparison_data_civil.docx")
                        output_comparison_pdf_file = os.path.join(downloads_dir, "comparison_data_civil.pdf")
                        scrape_comparison(url_comparison_civil, class_name, output_comparison_word_file,output_comparison_pdf_file)
                        response_text = f"Comparison data has been downloaded for {command} as a Word and PDF file."

                    elif "comparison for electrical engineering" in command:
                        # Extract and save the comparison table                        
                        url_comparison_electrical = "https://collegedunia.com/college-compare?entity_1=28749,1938&entity_2=25997,1938&entity_3=25503,1938&entity_4=25914,1938&popup=true"
                        output_comparison_word_file = os.path.join(downloads_dir, "comparison_data_electrical.docx")
                        output_comparison_pdf_file = os.path.join(downloads_dir, "comparison_data_electrical.pdf")
                        scrape_comparison(url_comparison_electrical, class_name, output_comparison_word_file,output_comparison_pdf_file)
                        response_text = f"Comparison data has been downloaded for {command} as a Word and PDF file."

                    # Check if the user's command contains "fees" and a course name
                    elif "fees" in command or "fee" in command and any(course in command for course, _, _ in course_data):
                        # Extract the course name from the command
                        cleaned_command = next((course for course, _, _ in course_data if course in command), None)
                        if cleaned_command:
                            response_text = f"The fee for {cleaned_command} is {course_data[[course for course, _, _ in course_data].index(cleaned_command)][1]}."
                        else:
                            response_text = "I'm sorry, I didn't understand your request."
                    # Check if the user's command contains "cutoff" and a course name
                    elif "cutoff" in command or "cut off" in command and any(course in command for course, _ in cutoff_data):
                        # Extract the course name from the command
                        cleaned_command = next((course for course, _ in cutoff_data if course in command), None)
                        if cleaned_command:
                            response_text = f"The cutoff for {cleaned_command} is {cutoff_data[[course for course, _ in cutoff_data].index(cleaned_command)][1]}."
                        else:
                            response_text = "I'm sorry, I didn't understand your request."
                    # Check if the user's command is "ranking"
                    elif "ranking" in command and any("pdeu college" or "pdpu college" or course in command for course, _ in cutoff_data):
                        if ranking_text:
                            response_text = f"The ranking of the university is {ranking_text}."
                        else:
                            response_text = "I'm sorry, the ranking data is not available."
                    else:
                        response_text = "I'm sorry, I didn't understand your request."
                else:
                    response_text = "I'm sorry, the data has not been scraped. Please enable data scraping by saying 'jarvis' again."

            # Print the recognized command and respond
            print("You said:", command)
            print("Jarvis:", response_text)

            # Speak the response
            speak(response_text)

        except sr.UnknownValueError:
            print("Sorry, I couldn't understand what you said.")
        except sr.RequestError as e:
            print("Sorry, there was an error connecting to the Google API. {0}".format(e))

if __name__ == "__main__":
    main()