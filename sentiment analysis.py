import pdfplumber
import csv
import os
import pandas as pd
import re
from collections import Counter
import shutil
import contractions


def load_lm_dictionary(file_path): #load loughran mcdonald dictionary
    lm_dict = pd.read_csv(file_path)
    sentiment_dict = {
        'positive': set(lm_dict[lm_dict['Positive'] != 0]['Word'].str.lower()),
        'negative': set(lm_dict[lm_dict['Negative'] != 0]['Word'].str.lower()),
        'uncertainty': set(lm_dict[lm_dict['Uncertainty'] != 0]['Word'].str.lower()),
        'litigious': set(lm_dict[lm_dict['Litigious'] != 0]['Word'].str.lower()),
        'strong_modal': set(lm_dict[lm_dict['Strong_Modal'] != 0]['Word'].str.lower()),
        'weak_modal': set(lm_dict[lm_dict['Weak_Modal'] != 0]['Word'].str.lower()),
        'constraining': set(lm_dict[lm_dict['Constraining'] != 0]['Word'].str.lower())
    }
    return sentiment_dict

def extract_participants(input_folder):
    try:
        with pdfplumber.open(input_folder) as pdf:
            if len(pdf.pages) < 2:
                return None, "", ""
            
            participants = {'EXECUTIVES': [], 'ANALYSTS': []}
            md_text = ""
            qa_text = ""
            presentation_started = False
            qa_started = False
            
            # Extract participants from the first page
            first_page = pdf.pages[0]
            extract_participants_from_page(first_page, participants)
            
            # Extract md_text and qa_text from the second page 
            for page in pdf.pages[1:]:
                bold_text = page.filter(lambda obj: obj["object_type"] == "char" and "Bold" in obj["fontname"])
                if bold_text is None:
                    continue
                
                bold_text = bold_text.extract_text()
                all_text = page.extract_text()
                
                bold_lines = bold_text.split('\n')
                all_lines = all_text.split('\n')
                
                for bold_line in bold_lines:
                    if "Presentation" in bold_line:
                        presentation_started = True
                        continue
                    if "Question" in bold_line and "Answer" in bold_line:
                        qa_started = True
                        continue
                    
                if presentation_started and not qa_started:
                    md_text += all_text + "\n"
                elif qa_started:
                    qa_text += all_text + "\n"
            
            # Clean up the extracted text
            md_text = md_text.strip()
            qa_text = qa_text.strip()
            
            return participants, md_text, qa_text
    except Exception as e:
        print(f"Error processing PDF: {str(e)}")
        return None, "", ""
    
def extract_participants_from_page(page, participants):
    try:
        # Extract bold and regular text
        bold_text = page.filter(lambda obj: obj["object_type"] == "char" and "Bold" in obj["fontname"])
        
        if bold_text is None:
            return
        
        bold_text = bold_text.extract_text()
        all_text = page.extract_text()
        
        if not bold_text or not all_text:
            return
        
        current_section = None
        # Clean text and remove extra spaces
        def clean_text(text):
            text = text.strip().lower()
            text = re.sub(r'[,&.\']', '', text)
            text = re.sub(r'\s+', ' ', text)  # Replace multiple spaces with single space
            return text.strip()
            
        lines = [clean_text(line) for line in all_text.split('\n')]
        current_name = None
        current_title = []
        
        # Clean bold text lines
        bold_lines = [clean_text(line) for line in bold_text.split('\n')]
        
        # Remove copyright lines
        lines = [line for line in lines if not line.startswith('copyright') and  
                not line.startswith('spglobal')]
        
        for i, line in enumerate(lines):
            if line == 'executives':
                if current_name and current_section and not any(name == current_name for name, _ in participants[current_section]):
                    participants[current_section].append((current_name, ' '.join(current_title).strip()))
                current_section = 'EXECUTIVES'
                current_name = None
                current_title = []
                continue
            elif line == 'analysts':
                if current_name and current_section and not any(name == current_name for name, _ in participants[current_section]):
                    participants[current_section].append((current_name, ' '.join(current_title).strip()))
                current_section = 'ANALYSTS'
                current_name = None
                current_title = []
                continue
            elif not line or line == 'call participants':
                continue
            
            if current_section:
                if any(name == line for name in bold_lines):
                    if current_name and not any(name == current_name for name, _ in participants[current_section]):
                        participants[current_section].append((current_name, ' '.join(current_title).strip()))
                    current_name = line
                    current_title = []
                elif current_name and line:
                    if not line.startswith('copyright'):
                        current_title.append(line)
        
        if current_name and current_section and not any(name == current_name for name, _ in participants[current_section]):
            participants[current_section].append((current_name, ' '.join(current_title).strip()))
            
    except Exception as e:
        print(f"Error extracting participants: {str(e)}")


def clean_text(md_text, qa_text):
    # Define the patterns
    pattern1 = r"Copyright Ãƒâ€šÃ‚Â© \d{4} S&P Global Market Intelligence, a division of S&P Global Inc\. All Rights reserved\.\nspglobal\.com/marketintelligence \d+\n.*\n"
    pattern2 = r"Copyright Ãƒâ€šÃ‚Â© \d{4} S&P Global Market Intelligence, a division of S&P Global Inc\. All Rights reserved\.\nspglobal\.com/marketintelligence \d+$"
    pattern3 = r".*EARNINGS CALL.*\d{4}\n"
    pattern4 = r"Presentation\n"
    pattern5 = r"Question and Answer\n"
    
    def clean_for_sentiment(text):
        expanded_text = contractions.fix(text)
        # Remove text within brackets and parentheses along with the brackets/parentheses themselves
        text = re.sub(r'\[.*?\]', '', expanded_text)
        text = re.sub(r'\(.*?\)', '', text)
        text = re.sub(r'\<.*?\>', '', text)
        # Remove specific characters while preserving line breaks
        text = text.replace(',', '')
        text = text.replace("'", '')
        text = text.replace('"', '')
        text = text.replace('%', '')
        text = text.replace('--', ' ')
        text = text.replace('-', ' ')
        text = text.replace('?', '')
        text = text.replace('[', '')
        text = text.replace(']', '')
        text = text.replace('(', '')
        text = text.replace(')', '')
        # Remove special characters while preserving line breaks
        text = re.sub(r'[^\w\s\n]', ' ', text)
        # Clean up multiple spaces while preserving line breaks
        text = re.sub(r' +', ' ', text)
        # Convert to lowercase
        text = text.lower()
        # Clean up multiple newlines but preserve single ones
        text = re.sub(r'\n\s*\n', '\n', text)
        
        return text.strip()

    # Clean md_text
    md_text = re.sub(pattern1, "", md_text)
    md_text = re.sub(pattern2, "", md_text)
    md_text = re.sub(pattern3, "", md_text)
    md_text = re.sub(pattern4, "", md_text)
    md_text = clean_for_sentiment(md_text)
    
    # Clean qa_text
    qa_text = re.sub(pattern1, "", qa_text)
    qa_text = re.sub(pattern2, "", qa_text)
    qa_text = re.sub(pattern3, "", qa_text)
    qa_text = re.sub(pattern5, "", qa_text)
    qa_text = clean_for_sentiment(qa_text)
    
    return md_text, qa_text

def detect_speakers_with_sentiment(text, participants, sentiment_dict):
    all_participants = {}
    speaker_sentiment = {}
    current_interaction = 0
    current_analyst = None
    current_executive = None

    for category in participants:
        for name, title in participants[category]:
            all_participants[name] = title
            speaker_sentiment[name] = []

    def analyze_sentiment(text, speaker, analyst=None):
        sentiment = {
            'positive': 0,
            'negative': 0,
            'uncertainty': 0,
            'litigious': 0,
            'strong_modal': 0,
            'weak_modal': 0,
            'constraining': 0,
            'word_count': 0,
            'text': [],
            'interaction': current_interaction,
            'analyst': analyst
        }
        
        words = [w for w in text.lower().split() if w]
        sentiment['word_count'] = len(words)
        
        # Use Counter to count occurrences of each word
        word_counts = Counter(words)
        
        for word, count in word_counts.items():
            for sentiment_type, word_set in sentiment_dict.items():
                if word in word_set:
                    sentiment[sentiment_type] += count
        
        return sentiment

    lines = text.split('\n')
    current_text = ""
    
    for line in lines:
        line = line.strip()
        if not line:
            continue
        
        speaker_found = False
        for participant_name in all_participants:
            if (line.lower().startswith(participant_name.lower()) or 
                (line.lower().startswith('operator') and participant_name.lower() in line.lower())):
                if current_text:
                    if current_executive:
                        sentiment = analyze_sentiment(current_text, current_executive, current_analyst)
                        speaker_sentiment[current_executive].append(sentiment)
                    elif current_analyst:
                        sentiment = analyze_sentiment(current_text, current_analyst)
                        speaker_sentiment[current_analyst].append(sentiment)
                
                current_text = ""
                if participant_name in [name for name, _ in participants['ANALYSTS']]:
                    current_analyst = participant_name
                    current_executive = None
                    current_interaction += 1
                else:
                    current_executive = participant_name
                
                speaker_found = True
                break
        
        if not speaker_found:
            current_text += " " + line

    # Process the last interaction
    if current_text:
        if current_executive:
            sentiment = analyze_sentiment(current_text, current_executive, current_analyst)
            speaker_sentiment[current_executive].append(sentiment)
        elif current_analyst:
            sentiment = analyze_sentiment(current_text, current_analyst)
            speaker_sentiment[current_analyst].append(sentiment)

    return speaker_sentiment

def detect_qa_interactions(qa_text, participants):
    interaction_count = 1
    current_analyst = None
    executive_analyst_pairs = {}
    current_executive = None
    word_count = 0
    lines = qa_text.split('\n')
    
    for line in lines:
        line = line.strip().lower()
        if not line:
            continue
        
        if 'operator' in line:
            # Only increment interaction count if a new analyst is introduced
            if ':' in line:
                next_analyst = line.split(':')[1].strip()
                if next_analyst != current_analyst:
                    interaction_count += 1
                    current_analyst = next_analyst
            current_executive = None
            word_count = 0
            continue
        
        speaker_found = False
        for speaker, _ in participants['EXECUTIVES'] + participants['ANALYSTS']:
            if line.startswith(speaker.lower()):
                if speaker in [name for name, _ in participants['ANALYSTS']]:
                    if current_analyst != speaker:
                        interaction_count += 1
                        current_analyst = speaker
                elif speaker in [name for name, _ in participants['EXECUTIVES']]:
                    current_executive = speaker
                word_count = 0
                speaker_found = True
                line = line[len(speaker):].strip()
                break
        
        if not speaker_found and current_executive and current_analyst:
            words = line.split()
            word_count += len(words)
            pair = (current_executive, current_analyst)
            if pair not in executive_analyst_pairs:
                executive_analyst_pairs[pair] = {'interaction': interaction_count, 'word_count': 0}
            executive_analyst_pairs[pair]['word_count'] += len(words)
    
    return interaction_count, executive_analyst_pairs


def extract_file_metadata(filename):
    # Remove 'processed_' if present
    clean_name = filename.replace('processed_', '')

    # Check if the filename is in the format 'processed_companyname_year'
    if re.match(r'^[^,]+_\d{4}$', clean_name):
        company_name, year = clean_name.rsplit('_', 1)
        quarter = '4'  # Assume quarter 4
    else:
        # Extract company name (everything before first comma)
        company_name = clean_name.split(',')[0].strip()

        # Extract quarter and year using regex
        quarter_match = re.search(r'Q(\d)\s+(\d{4})', clean_name)
        if quarter_match:
            quarter = quarter_match.group(1)
            year = quarter_match.group(2)
        else:
            quarter = ""
            year = ""

    return company_name, quarter, year

def save_sentiment_analysis(md_sentiment, qa_sentiment, participants, output_folder, filename, qa_text, output_file):
    company_name, quarter, year = extract_file_metadata(filename)
    interaction_counts, executive_analyst_pairs = detect_qa_interactions(qa_text, participants)
    roles = {}
    for name, _ in participants['EXECUTIVES']:
        roles[name] = 'Executive'
    for name, _ in participants['ANALYSTS']:
        roles[name] = 'Analyst'
    titles = {}
    for category in participants:
        for name, title in participants[category]:
            titles[name] = title
    
    with open(output_file, mode='a', newline='', encoding='utf-8') as file:
        writer = csv.writer(file)
        
        # Write MD section data
        for speaker, data_list in md_sentiment.items():
            for data in data_list:
                writer.writerow([
                    company_name, quarter, year,
                    speaker, roles.get(speaker, ""), titles.get(speaker, ""),
                    "Presentation", data["word_count"], data["positive"], data["negative"],
                    data["uncertainty"], data["litigious"], data["strong_modal"],
                    data["weak_modal"], data["constraining"], 0, ""
                ])
        
        # Write QA section data
        current_analyst = None
        for speaker, data_list in qa_sentiment.items():
            for data in data_list:
                if roles.get(speaker) == 'Analyst':
                    current_analyst = speaker
                writer.writerow([
                    company_name, quarter, year,
                    speaker, roles.get(speaker, ""), titles.get(speaker, ""),
                    "Q&A", data["word_count"], data["positive"], data["negative"],
                    data["uncertainty"], data["litigious"], data["strong_modal"],
                    data["weak_modal"], data["constraining"], data['interaction'],
                    current_analyst if roles.get(speaker) == 'Executive' else ""
                ])

    return output_file

def process_pdf(input_path, output_file):
    try:
        sentiment_dict = load_lm_dictionary(lm_dict_path)
        participants, md_text, qa_text = extract_participants(input_path)
        if participants is None:
            raise Exception("Failed to extract participants")
        cleaned_md_text, cleaned_qa_text = clean_text(md_text, qa_text)
        md_speaker_sentiment = detect_speakers_with_sentiment(cleaned_md_text, participants, sentiment_dict)
        qa_speaker_sentiment = detect_speakers_with_sentiment(cleaned_qa_text, participants, sentiment_dict)
        filename = os.path.basename(input_path)
        save_sentiment_analysis(md_speaker_sentiment, qa_speaker_sentiment, participants, output_folder, filename, cleaned_qa_text, output_file)
        return True
    except Exception as e:
        print(f"Error processing {os.path.basename(input_path)}: {str(e)}")
        return False

# Example usage
desktop_path = os.path.expanduser("~/Desktop")
input_folder = os.path.join(desktop_path, "call")
output_folder = os.path.join(desktop_path, "results")
lm_dict_path = os.path.join(desktop_path, "dictionary", "Loughran-McDonald_MasterDictionary_1993-2023.csv")
# Create output folder if it doesn't exist
os.makedirs(output_folder, exist_ok=True)

def process_folder(input_folder):
    error_folder = os.path.join(output_folder, "error_pdfs")
    os.makedirs(error_folder, exist_ok=True)
    
    output_file = os.path.join(output_folder, "speaker_sentiment_analysis.csv")
    
    # Write header to the output file
    with open(output_file, mode='w', newline='', encoding='utf-8') as file:
        writer = csv.writer(file)
        writer.writerow([
            "Company", "Quarter", "Year", "Speaker", "Role", "Title", "Section",
            "Word Count", "Positive", "Negative", "Uncertainty", "Litigious",
            "Strong Modal", "Weak Modal", "Constraining", "Interaction", "Analyst"
        ])
    
    for filename in os.listdir(input_folder):
        if filename.lower().endswith('.pdf'):
            input_path = os.path.join(input_folder, filename)
            success = process_pdf(input_path, output_file)
            if not success:
                try:
                    error_file_path = os.path.join(error_folder, filename)
                    shutil.move(input_path, error_file_path)
                    print(f"Moved {filename} to error folder")
                except Exception as e:
                    print(f"Failed to move {filename} to error folder: {str(e)}")

# Run the processing
if __name__ == "__main__":
    process_folder(input_folder)
