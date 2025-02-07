import google.generativeai as genai
import time
from datetime import datetime
import json
from global_signals import global_signals
import re
import os
import PyPDF2
from docx import Document
import ast

def read_pdf(file_path):
    """Reads and returns text from a PDF file."""
    text = ""
    try:
        with open(file_path, 'rb') as file:
            reader = PyPDF2.PdfReader(file)
            for page in reader.pages:
                text += page.extract_text() + "\n"
    except Exception as e:
        print(f"Error reading PDF {file_path}: {e}")
    return text

def read_docx(file_path):
    """Reads and returns text from a DOCX file."""
    text = ""
    try:
        doc = Document(file_path)
        for paragraph in doc.paragraphs:
            text += paragraph.text + "\n"
    except Exception as e:
        print(f"Error reading DOCX {file_path}: {e}")
    return text

def _extract_list_from_string(text):
    """
    Safely extracts a Python list from a string and returns it as a *string*
    representation suitable for JSON, handling various Gemini output quirks.
    """
    match = re.search(r'\[[^\]]*\]', text)  # Find the first list-like structure
    if match:
        list_str = match.group(0)
        try:
            # Use ast.literal_eval for safe evaluation
            parsed_list = ast.literal_eval(list_str)
            if isinstance(parsed_list, list):
                # Convert the list to a JSON-compatible string representation
                return json.dumps(parsed_list)
        except (SyntaxError, ValueError):
            pass  # Fallthrough to return None if parsing fails
    return '[]'  # Return empty list *string* if no valid list found

max_wait_time = 200

# --- REVISED PROMPTS ---
prompts = {
    'prompt2_firstimpr': (
        """You're an Assessor at ORMIT Talent.  Give a concise first impression of a trainee (max 40 words).
Focus on: Overall vibe, speech, body language, and emotional tone.
Don't judge: Rely *only* on assessor observations in 'Assessment Notes'.
Output: One short paragraph (max 35 words).  *Only* the first impression, no extra words or formatting.
"""
    ),
    "prompt3_personality": (
        """You're an Assessor at ORMIT Talent. Describe the trainee's personality (250-400 words).
End with:
  1. PAPI test summary (50-100 words, *must* mention "PAPI test").
  2. Very short final summary (2-3 sentences).

Use: 'Context and Task description', 'Assessment Notes', 'PAPI Feedback'.
'Personality Section Examples' is for *structure only*, not writing style.

Instructions:
  - Be conversational, professional.
  - Identify main traits, strengths, weaknesses.  Avoid technical skills.
  - Give *examples* from 'Assessment Notes' for traits.
  - Be balanced (strengths and areas for improvement).
  - Frame improvements as learning opportunities.
  - Use simple language, be realistic.

Output:  *Only* the personality description, PAPI summary, and final summary.  No extra text or formatting.  Separate the three sections with blank lines. Do *not* include any labels or section titles.
"""
    ),
     'prompt4_cogcap_scores': (
        """Read 'Context and Task Description' and 'Capacity test results'.
Find the six cognitive capacity scores: general ability, speed, accuracy, verbal, numerical, abstract.

Output: A *string* containing a Python list.  *Nothing else*.
Order: [general_ability, speed, accuracy, verbal, numerical, abstract]

Example: "[75, 80, 85, 70, 65, 78]"

The output *must* be a directly usable Python list string (enclosed in double quotes). No extra text, no backticks, no "python" labels.
"""
    ),
    'prompt4_cogcap_remarks': (
        """Read 'Capacity test results'.
Write a 2-3 sentence summary interpreting the results.
Focus on:
  - Overall general ability.
  - Speed vs. accuracy.
  - Sub-test performance (verbal, numerical, abstract).

Output: *Only* the summary text.  No labels, formatting, or extra sentences.
"""
    ),
    'prompt5_language': (
        """Determine the trainee's language levels (Dutch, French, English).
Use: 'Context and Task description' and 'Assessment Notes'.

Instructions:
  1. If 'Assessment Notes' specifies levels (e.g., 'B2'), use those.
  2. Otherwise, use the guide in 'Context and Task description' (section '5. Language Skills').

Output: A *string* containing a Python list: [Dutch level, French level, English level]
Example: "['C1', 'B2', 'C2']"

*Only* the list string. No other text.  The output *must* be a directly usable Python list string (enclosed in double quotes). No backticks.
"""
    ),
    'prompt6a_conqual': (
        """Identify 6-7 of the trainee's *strengths*.
Use: 'Context and Task Description', 'Assessment Notes', 'PAPI Feedback'.

Instructions:
  - Short, practical descriptions (under 10 words each).
  - Simple language.

Output: A *string* containing a Python list.
Example: "['Good listener', 'Communicates clearly', 'Works well in teams']"

*Only* the list string.  No other text. The output *must* be a directly usable Python list string (enclosed in double quotes).  No backticks.
"""
    ),
    'prompt6b_conimprov': (
        """Identify 4-5 of the trainee's *improvement points*.
Use: 'Context and Task Description', 'Assessment Notes', 'PAPI Feedback'.

Instructions:
  - Short, practical descriptions (under 10 words each).
  - Simple language.

Output: A *string* containing a Python list.
Example: "['Needs more assertiveness', 'Can be more proactive']"

*Only* the list string. No other text. The output *must* be a directly usable Python list string (enclosed in double quotes). No backticks.
"""
    ),
   'prompt7_qualscore': (
        """Match the trainee's qualities to the *green-highlighted* descriptions in the 'MCP profile'.
Create a scored list (20 numbers: 0s and 1s).

Output: A *string* containing a Python list.
Example: "[0, 1, 0, 1, 1, 0, 0, 1, 0, 0, 1, 0, 0, 0, 0, 0, 1, 0, 0, 0]"

*Only* the list string.  No other text. The output *must* be a directly usable Python list string (enclosed in double quotes). No backticks.
"""
    ),
    'prompt7_qualscore_data': (
        """Match trainee's qualities to *green-highlighted* descriptions in 'The Data Chiefs profile'.
Create a scored list (23 numbers: 0s and 1s).

Use: 'Assessment Notes', and 'The Data Chiefs profile'.
Scoring:
  1: Strong fit.
  0: Relevant, but not a top strength.
  -1: (Don't use -1 here, only 0 or 1).

Output: A *string* containing a Python list.
Example: "[0, 1, 0, 1, 1, 0, 0, 1, 0, 0, 1, 0, 0, 0, 0, 0, 1, 0, 0, 0, 0, 0, 0]"

*Only* the list string.  No other text. The output *must* be a directly usable Python list string (enclosed in double quotes).  No backticks.
"""
    ),
    'prompt8_datatools': (
        """Analyze 'Assessment Notes' for data skill proficiency.
Output a *string* containing a Python list of 5 numbers (-1, 0, or 1).

Skills (in order):
  1. Excel/VBA
  2. Power BI/Tableau/Qlik Sense
  3. Python/R
  4. SQL
  5. Azure Databricks

Scale:
  -1: Beginner/Improvement point.
  0: Average
  1: Proficient

Output:  *Only* the list string.
Example: "[-1, 1, 0, 1, -1]"

The output *must* be a directly usable Python list string (enclosed in double quotes). No extra text, no backticks.
"""
    ),
    'prompt9_interests': (
        """Identify 3-5 data-related interests from 'Assessment Notes'.

Output: A *string* containing a Python list.
Example: "['Machine Learning', 'Data Visualization']"

*Only* the list string.  No other text. The output *must* be a directly usable Python list string (enclosed in double quotes). No backticks.
"""
    ),
}

def send_prompts(data):
    print('Prompting started')
    global_signals.update_message.emit("Connecting to Gemini...")

    # Gemini Pro API setup
    GOOGLE_API_KEY = data["Gemini Key"]
    genai.configure(api_key=GOOGLE_API_KEY)
    model = genai.GenerativeModel(model_name="gemini-2.0-flash-001") # Or your preferred, compatible model.

    # Filename setup
    current_time = datetime.now()
    formatted_time = current_time.strftime("%m%d%H%M")
    appl_name = data["Applicant Name"]
    filename_with_timestamp = f"{appl_name}_{formatted_time}.json"

    # File paths
    path_to_notes = r'temp/Assessment Notes.pdf'
    path_to_persontest = r'temp/PAPI Feedback.pdf'
    path_to_cogcap = r'temp/Cog. Test.pdf'
    path_to_contextfile = r'resources/Context and Task Description.docx'
    path_to_toneofvoice = r'resources/Examples Personality Section.docx'
    path_to_mcpprofile = r'resources/The MCP Profile.docx'
    path_to_dataprofile = r'resources/The Data Chiefs profile.docx'

    lst_files = [
        path_to_notes,
        path_to_persontest,
        path_to_cogcap,
        path_to_contextfile,
        path_to_toneofvoice,
    ]

    selected_program = data["Traineeship"]
    if selected_program == 'DATA':
        lst_files.append(path_to_dataprofile)
    else:
        lst_files.append(path_to_mcpprofile)

    # Pre-load file contents
    file_contents = {}
    for file_path in lst_files:
        file_name = os.path.basename(file_path)
        if file_path.endswith('.pdf'):
            file_contents[file_name] = read_pdf(file_path)
        elif file_path.endswith('.docx'):
            file_contents[file_name] = read_docx(file_path)
        else:
            print(f"Warning: Unsupported file type: {file_path}")
            file_contents[file_name] = ""

    global_signals.update_message.emit("Files uploaded, starting prompts...")

    # Determine which prompts to use based on program
    prompt_keys = [
        'prompt2_firstimpr', 'prompt3_personality', 'prompt4_cogcap_scores',
        'prompt4_cogcap_remarks', 'prompt5_language', 'prompt6a_conqual',
        'prompt6b_conimprov', 'prompt7_qualscore_data', 'prompt8_datatools',
        'prompt9_interests'
    ]  # Data Chiefs prompts

    if selected_program == 'MCP':
        prompt_keys = [
            'prompt2_firstimpr', 'prompt3_personality', 'prompt4_cogcap_scores',
            'prompt4_cogcap_remarks', 'prompt5_language', 'prompt6a_conqual',
            'prompt6b_conimprov', 'prompt7_qualscore'
        ]

    lst_prompts = prompt_keys
    print(lst_prompts)

    # Prompts requiring list output
    list_output_prompts = [
        'prompt4_cogcap_scores', 'prompt6a_conqual', 'prompt6b_conimprov',
        'prompt7_qualscore', 'prompt7_qualscore_data', 'prompt8_datatools',
        'prompt9_interests', 'prompt5_language'
    ]

    results = {}
    start_time_all = time.time()

    for promno, prom in enumerate(lst_prompts, start=1):
        print(prom)
        global_signals.update_message.emit(f"Submitting prompt {promno}/{len(lst_prompts)}, please wait...")

        prompt_text = prompts[prom]
        context = "\n\n---\n\n".join([f"File: {file_name}\nContent:\n{content}"
                                        for file_name, content in file_contents.items()])
        full_prompt = f"{prompt_text}\n\nUse the following files to answer:\n{context}"

        try:
            response = model.generate_content(full_prompt)
            output_text = response.text
            print(f"Prompt: {prom}")
            print(f"Raw Output: {output_text}")

            # --- Crucial Post-Processing ---
            if prom in list_output_prompts:
                results[prom] = _extract_list_from_string(output_text)
            else:
                results[prom] = output_text.strip()

        except Exception as e:
            print(f"Error processing prompt {prom}: {e}")
            results[prom] = ""  # Store empty string on error

        if time.time() - start_time_all > max_wait_time:
            print("Timeout for all prompts reached.")
            break

    # Save results to JSON
    with open(filename_with_timestamp, 'w') as json_file:
        json.dump(results, json_file, indent=4)

    global_signals.update_message.emit("Prompting finished, generating report...")
    return filename_with_timestamp
