import pdfplumber
import csv
import os
import pandas as pd
import re
from collections import Counter
import shutil

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
    pattern1 = r"Copyright Â© \d{4} S&P Global Market Intelligence, a division of S&P Global Inc\. All Rights reserved\.\nspglobal\.com/marketintelligence \d+\n.*\n"
    pattern2 = r"Copyright Â© \d{4} S&P Global Market Intelligence, a division of S&P Global Inc\. All Rights reserved\.\nspglobal\.com/marketintelligence \d+$"
    pattern3 = r".*EARNINGS CALL.*\d{4}\n"
    pattern4 = r"Presentation\n"
    pattern5 = r"Question and Answer\n"
    
    def clean_for_sentiment(text):
        # Remove text within brackets and parentheses along with the brackets/parentheses themselves
        text = re.sub(r'\[.*?\]', '', text)
        text = re.sub(r'\(.*?\)', '', text)
        text = re.sub(r'\<.*?\>', '', text)
        # Remove specific characters while preserving line breaks
        text = text.replace(',', '')
        text = text.replace("'", '')
        text = text.replace('"', '')
        text = text.replace('%', '')
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
    # Create dictionaries for participants and sentiment
    all_participants = {}
    speaker_sentiment = {}
    
    # Initialize dictionaries first
    for category in participants:
        for name, title in participants[category]:
            all_participants[name] = title
            speaker_sentiment[name] = {
                'positive': 0,
                'negative': 0,
                'uncertainty': 0,
                'litigious': 0,
                'strong_modal': 0,
                'weak_modal': 0,
                'constraining': 0,
                'word_count': 0,
                'text': []
            }
    
    def analyze_sentiment(text, speaker):
        # Split text into words and remove empty strings
        words = [w for w in text.lower().split() if w]
        # Add to total word count
        speaker_sentiment[speaker]['word_count'] += len(words)
        
        for word in words:
            for sentiment_type, word_set in sentiment_dict.items():
                if word in word_set:
                    speaker_sentiment[speaker][sentiment_type] += 1
    
    current_speaker = None
    lines = text.split('\n')
    
    for line in lines:
        line = line.strip()
        if not line:
            continue
            
        speaker_found = False
        for participant_name in all_participants:
            # Check both direct starts and mentions in operator introductions
            if (line.lower().startswith(participant_name.lower()) or 
                (line.lower().startswith('operator') and participant_name.lower() in line.lower())):
                current_speaker = participant_name
                # If operator is introducing, take the next line as the speaker's text
                if 'operator' in line.lower():
                    continue
                clean_line = line[len(participant_name):].strip()
                if clean_line:
                    speaker_sentiment[current_speaker]['text'].append(clean_line)
                    analyze_sentiment(clean_line, current_speaker)
                speaker_found = True
                break
        
        # Handle continuation of current speaker's text
        if not speaker_found and current_speaker and line:
            speaker_sentiment[current_speaker]['text'].append(line)
            analyze_sentiment(line, current_speaker)
    
    return speaker_sentiment

def detect_qa_interactions(qa_text, participants):
    interaction_count = {}
    current_interaction = 0
    lines = qa_text.split('\n')
    current_speakers = set()
    current_analyst = None
    executive_analyst_pairs = {}

    for line in lines:
        line = line.strip().lower()
        if not line:
            continue

        if 'operator' in line:
            current_interaction += 1
            current_speakers.clear()
            current_analyst = None

        for speaker, _ in participants['EXECUTIVES'] + participants['ANALYSTS']:
            if line.startswith(speaker.lower()):
                if speaker in [name for name, _ in participants['ANALYSTS']]:
                    current_analyst = speaker
                elif speaker in [name for name, _ in participants['EXECUTIVES']]:
                    if current_analyst:
                        pair = (speaker, current_analyst)
                        if pair not in executive_analyst_pairs:
                            executive_analyst_pairs[pair] = current_interaction
                interaction_count[speaker] = current_interaction
                current_speakers.add(speaker)

    return interaction_count, executive_analyst_pairs

def extract_file_metadata(filename):
    # Remove 'processed_' if present
    clean_name = filename.replace('processed_', '')
    
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

def save_sentiment_analysis(md_sentiment, qa_sentiment, participants, output_folder, filename, qa_text):
    os.makedirs(output_folder, exist_ok=True)
    output_file = os.path.join(output_folder, "speaker_sentiment_analysis.csv")
    
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
    
    with open(output_file, mode='w', newline='', encoding='utf-8') as file:
        writer = csv.writer(file)
        writer.writerow([
            "Company", "Quarter", "Year", "Speaker", "Role", "Title", "Section", "Word Count",
            "Positive", "Negative", "Uncertainty", "Litigious", "Strong Modal", "Weak Modal",
            "Constraining", "Interaction", "Analyst"
        ])
        
        # Write MD section data - exclude analysts
        for speaker, data in md_sentiment.items():
            if roles.get(speaker) != 'Analyst':
                writer.writerow([
                    company_name, quarter, year, speaker, roles.get(speaker, ""),
                    titles.get(speaker, ""), "Presentation", data["word_count"],
                    data["positive"], data["negative"], data["uncertainty"], data["litigious"],
                    data["strong_modal"], data["weak_modal"], data["constraining"],
                    0, ""  # No interactions or analysts in presentation
                ])
        
        # Write QA section data
        for speaker, data in qa_sentiment.items():
            if roles.get(speaker) == 'Analyst':
                writer.writerow([
                    company_name, quarter, year, speaker, roles.get(speaker, ""),
                    titles.get(speaker, ""), "Q&A", data["word_count"],
                    data["positive"], data["negative"], data["uncertainty"], data["litigious"],
                    data["strong_modal"], data["weak_modal"], data["constraining"],
                    interaction_counts.get(speaker, 0), ""
                ])
            else:  # Executive
                for (exec_name, analyst_name), interaction in executive_analyst_pairs.items():
                    if exec_name == speaker:
                        writer.writerow([
                            company_name, quarter, year, speaker, roles.get(speaker, ""),
                            titles.get(speaker, ""), "Q&A", data["word_count"],
                            data["positive"], data["negative"], data["uncertainty"], data["litigious"],
                            data["strong_modal"], data["weak_modal"], data["constraining"],
                            interaction, analyst_name
                        ])

    return output_file

def process_pdf(input_path):
    try:
        # Load the sentiment dictionary
        sentiment_dict = load_lm_dictionary(lm_dict_path)
        
        # Extract and clean text
        participants, md_text, qa_text = extract_participants(input_path)
        if participants is None:
            raise Exception("Failed to extract participants")
            
        cleaned_md_text, cleaned_qa_text = clean_text(md_text, qa_text)
        
        # Perform sentiment analysis
        md_speaker_sentiment = detect_speakers_with_sentiment(cleaned_md_text, participants, sentiment_dict)
        qa_speaker_sentiment = detect_speakers_with_sentiment(cleaned_qa_text, participants, sentiment_dict)
        
        # Get filename for metadata
        filename = os.path.basename(input_path)
        
        # Save results with metadata and qa_text
        output_file = save_sentiment_analysis(md_speaker_sentiment, qa_speaker_sentiment, 
                                            participants, output_folder, filename, cleaned_qa_text)
        return True
        
    except Exception as e:
        print(f"Error processing {os.path.basename(input_path)}: {str(e)}")
        return False


# Example usage
input_folder = os.path.expanduser("~/Desktop/call")
output_folder = os.path.expanduser("~/Desktop/results")
lm_dict_path = os.path.expanduser("~/Desktop/dictionary/Loughran-McDonald_MasterDictionary_1993-2023.csv")

# Process all PDFs in the input folder
def process_folder(input_folder):
    # Create error folder if it doesn't exist
    error_folder = os.path.join(output_folder, "error_pdfs")
    os.makedirs(error_folder, exist_ok=True)
    
    for filename in os.listdir(input_folder):
        if filename.lower().endswith('.pdf'):
            input_path = os.path.join(input_folder, filename)
            # Process the PDF and check if it was successful
            success = process_pdf(input_path)
            # If processing failed, move the file to error folder
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
