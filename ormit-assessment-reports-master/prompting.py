import time
from openai import OpenAI
from datetime import datetime
import json
from global_signals import global_signals

def get_custom_key_list(prompts):
    custom_keys = []
    keys_to_go_through = list(prompts.keys())
    for k in prompts.keys():
        custom_keys.append(k)  # Add the current key
    return custom_keys  

max_wait_time = 150  # seconds
prompts = {
    'prompt2_firstimpr': (
        "**You're a Assessor at ORMIT Talent.**  \n"
        "**Your task:** Give a quick, honest first impression of a trainee (max 40 words).  What were they like when you first met them?\n"
        "**Look at:** 'Assessment Notes Candidate' document.\n"
        "**Focus on:**  Their body language, how they spoke, and just their general vibe.\n"
        "**Don't judge:** Just say what you saw and heard in the first few minutes.\n"
        "**Output:** One short paragraph (max 40 words).  Just the first impression, nothing else. **Only a single line of text with the first impression, no extra words or formatting.**"
    ),
    "prompt3_personality": (
        "**You're a Assessor at ORMIT Talent.**\n"
        "**Your task:**  Describe the personality of a trainee named 'Piet' (around 350-500 words). Include a summary of his PAPI personality test and a short final summary.\n"
        "**Use these documents:** 'Context & Task', 'Assessment Notes Candidate', 'Personality Test Results'.  Look at 'Personality Section Examples' ONLY to see how it's structured, not for the writing style. We want simpler language.\n"
        "**Instructions:**\n"
        "1. **Keep it real & Easy to Read:** Write like you're talking to a colleague.  Be down-to-earth, maybe a little bit funny, but still professional.  Check 'Personality Section Examples' to get the structure right, but use simpler words.\n"
        "2. **Figure out Piet's Personality:** What are his main personality traits? What are his strengths and weaknesses? Think about how he acts with others, in teams, and in important moments from the 'Assessment Notes'.\n"
        "3. **Show, Don't Just Tell (Give Examples):** For each personality trait you talk about, give a specific example from the 'Assessment Notes Candidate'.  Explain *when* and *how* Piet showed this trait, describe what happened.\n"
        "4. **See Both Sides:**  Point out Piet's strengths, but also talk honestly and helpfully about things he can improve.  Frame areas for improvement as chances to learn and grow.\n"
        "5. **Simple Words & Tone:** Don't use fancy language or get too excited. Keep it realistic and grounded.\n"
        "6. **Word Count for Personality Snapshot:** The main personality description (before summaries) should be about 350-500 words.\n"
        "7. **PAPI Test Summary Block:** After the main description, write a separate short paragraph summarizing the 'Personality Test Results'. **Make sure to mention 'PAPI test' in this summary.** Keep it short (around 50-100 words).\n"
        "8. **Final Short Summary:**  Finish with a very brief summary of Piet's personality (2-3 sentences). This comes *after* the PAPI summary and wraps up the main points.\n"
        "**Output:**  The output should be just the text, like this (no extra labels or formatting):\n"
        "   [Personality Snapshot Text (350-500 words)]\n"
        "   [PAPI Test Summary Block (50-100 words, including 'PAPI' source)]\n"
        "   [Final Short Summary (2-3 sentences)]\n"
        "   **Make sure ONLY these three parts are in the output, with line breaks in between. No other text or labels.**"
    ),
    'prompt4_cogcap_scores': (
        "1. Read the 'Context and Task Description' file and 'Capacity test results' file carefully to understand what's needed.\n"
        "2. Find the cognitive capacity scores for the trainee. There are six scores: general ability, speed, accuracy, verbal, numerical, and abstract.\n"
        "3. Write down the scores in the exact order and format from the 'Context and Task Description' file.\n"
        "4. Give the output as a Python list of scores in the following format: [score_general_ability, score_speed, score_accuracy, score_verbal, score_numerical, score_abstract].\n"
        "5. Don't add any other words or text before the list, just the list.\n"
        "Output: Provide ONLY a valid Python list of the six cognitive capacity scores as requested in step 4, enclosed in a markdown code block. No extra labels, formatting, or introductory text around the code block, just the raw output:\n"
        "```python\n[score_general_ability, score_speed, score_accuracy, score_verbal, score_numerical, score_abstract]\n```"
    ),
    'prompt4_cogcap_remarks': (
        "1. Read the 'Capacity test results' file carefully to understand what's needed.\n"
        "2. Write a concise, 3-4 sentence summary interpreting the results. Focus on key strengths and relative performance (e.g., \"above average,\" \"average\") based on the percentile/stens provided. Do not explain scoring scales or define terms.\n"
        "3. Structure the summary to highlight:\n"
        "- Overall general ability (e.g., mental agility).\n"
        "- Speed vs. accuracy balance (e.g., fast pace with typical precision).\n"
        "- Sub-test performance (verbal, numerical, abstract) in simple terms. \n"
        "4. Frame the interpretation neutrally and professionally, as for a job assessment context.\n"
        "5. Output: Provide ONLY the summary in a markdown code block, no labels or formatting:\n"
    ),
    'prompt5_language': (
        "**You're a Assessor at ORMIT Talent.**\n"
        "**Your task:**  Figure out the trainee's language skills in Dutch, French, and English.\n"
        "**Use:** 'Context & Task' and 'Assessment Notes Candidate' documents.\n"
        "**Instructions:**\n"
        "1. Check the 'Assessment Notes'. If it says the trainee's language level (like 'B2' or 'C1'), use that.\n"
        "2. If not, use the guide in 'Context & Task' (section '5. Language Skills') to decide the level.\n"
        "3. Put the levels in a Python list, in this order: Dutch, French, English.\n"
        "**Output:** **ONLY** a Python list: [Dutch level, French level, English level]. **Nothing else. No words, no labels, no formatting, just the levels (like 'A2', 'C1', etc.) in a list.  It needs to be directly usable in Python code.**"
    ),
    'prompt6a_conqual': (
        "1. Read the 'Context and Task Description' file, Assessment Notes Candidate, and Personality Test Results document thoroughly to fully understand the task requirements.\n"
        "2. Identify and collect 6 or 7 of the trainee's strongest qualities based on the assessment notes. Each quality should be clear, down-to-earth, and in simple language. Focus on short, practical descriptions of skills or behaviors, avoiding complex or formal words.\n"
        "3. Keep each statement under 10 words, focusing on clear, everyday language.\n"
        "4. Provide the output as a Python list in the following format: [first_quality, second_quality, third_quality, fourth_quality, fifth_quality, sixth_quality, seventh_quality].\n"
        "5. Do not include any additional information, explanations, or text; return only the specified list."
    ),
    'prompt6b_conimprov': (
        "1. Read the 'Context and Task Description' file, Assessment Notes Candidate, and Personality Test Results document thoroughly to fully understand the task requirements.\n"
        "2. Identify and collect 4 or 5 of the trainee's improvement/development points based on the assessment notes. Each development point should be clear, down-to-earth, and in simple language. Focus on short, practical descriptions of skills or behaviors, avoiding complex or formal words.\n"
        "3. Keep each statement under 10 words, focusing on clear, everyday language.\n"
        "4. Provide the output as a Python list in the following format: [first_improvement, second_improvement, third_improvement, fourth_improvement, fifth_improvement].\n"
        "5. Do not include any additional information, explanations, or text; return only the specified list."
    ),
    'prompt7_qualscore': (
        "**You're a  Assessor at ORMIT Talent.**\n"
        "**Your task:** Look at the trainee's best qualities and match them to the **green-highlighted descriptions** in the MCP profile to make a scored list. This list shows the trainee's strengths based on the MCP profile.\n"
        "**Use:** 'Context & Task', 'Assessment Notes Candidate', and 'MCP profile' documents.\n"
        "**You already have:** A list of 6-7 of their best qualities.\n"
        "**Scoring:**\n"
        "* **1:** Really strong - The quality really fits the MCP profile description.\n"
        "* **0:** Good potential - The quality is relevant but not a top strength in the MCP profile area.\n"
        "* **-1:** Needs improvement - (Don't use this score for 'best qualities', only use 0 or 1 here).\n"
        "**Instructions:**\n"
        "1. Look at the MCP profile document, **only** at the descriptions under the **green** headings.\n"
        "2. For EACH best quality, find the **single best description** from the **green parts** of the MCP profile that matches.\n"
        "3. Make a Python list of **20 numbers**, all starting as **0s**. This list matches the 20 parts of the MCP profile in order.\n"
        "4. For each quality you matched, change the matching **0** in your list to a **1**.  If a quality doesn't really fit any green description, leave it as **0**.\n"
        "5. Make sure your list has between **5 and 7 ones**, for the 6-7 best qualities.\n"
        "**Output:** **ONLY** a Python list of 20 numbers (0s and 1s). **Nothing else.**  No words, no explanations, no labels, no formatting.  The output must be directly copy-pasteable and usable as a Python list in code. Example: `[0, 1, 0, 1, 1, 0, 0, 1, 0, 0, 1, 0, 0, 0, 0, 0, 1, 0, 0, 0]`"
    ),
    'prompt7_improvscore': (
        "**You're a Assessor at ORMIT Talent.**\n"
        "**Your task:** Look at the trainee's improvement areas and match them to the **green-highlighted descriptions** in the MCP profile to make a scored list. This list shows where the trainee can grow, based on the MCP profile.\n"
        "**Use:** 'Context & Task', 'Assessment Notes Candidate', and 'MCP profile' documents.\n"
        "**You already have:** A list of 3-5 areas for improvement.\n"
        "**Scoring:**\n"
        "* **1:** Really strong - (Don't use this score for 'improvement areas', only use 0 or -1 here).\n"
        "* **0:** Good potential - The improvement area is relevant but doesn't go against the MCP profile description.\n"
        "* **-1:** Needs improvement - The improvement area clearly shows they need to develop in this part of the MCP profile.\n"
        "**Instructions:**\n"
        "1. Look at the MCP profile document, **only** at the descriptions under the **green** headings.\n"
        "2. For EACH improvement area, find the **single best description** from the **green parts** of the MCP profile that matches.\n"
        "3. Make a Python list of **20 numbers**, all starting as **0s**. This list matches the 20 parts of the MCP profile in order.\n"
        "4. For each improvement area you matched, change the matching **0** in your list to a **-1**. If an improvement area doesn't really relate to a green description, leave it as **0**.\n"
        "5. Make sure your list has between **3 and 5 negative ones (-1s)**, for the 3-5 improvement areas.\n"
        "**Output:** **ONLY** a Python list of 20 numbers (0s and -1s). **Nothing else.** No words, no explanations, no labels, no formatting. The output must be directly copy-pasteable and usable as a Python list in code. Example: `[0, -1, 0, 0, -1, 0, 0, 0, 0, 0, 0, -1, 0, 0, 0, 0, 0, 0, 0, 0]`"
    )
}

def send_prompts(data):
    print('Prompting started')
    global_signals.update_message.emit("Connecting to OpenAI...")
    results = {}
    #For filename:
    current_time = datetime.now()
    formatted_time = current_time.strftime("%m%d%H%M")  # Format as MMDDHHMinMin

    #Custom: Redacted anonymous files
    path_to_notes = 'temp/Assessment Notes.pdf'
    path_to_persontest = 'temp/PAPI Feedback.pdf'
    path_to_cogcap = 'temp/Cog. Test.pdf'
    #Default

    path_to_contextfile = r'resources\Context and Task Description.docx'
    path_to_toneofvoice = r'resources\Examples Personality Section.docx'
    path_to_mcpprofile = r'resources\The MCP Profile.docx'
    
    start_time = time.time()
    
    lst_files = [path_to_notes,
                 path_to_persontest,
                 path_to_cogcap,
                 path_to_contextfile,
                 path_to_toneofvoice,
                 path_to_mcpprofile
                 ]
    
    mykey = data["OpenAI Key"]
    
    # Initialize the OpenAI client with your API key
    client = OpenAI(api_key=mykey)
    
    # Create the assistant
    assistant = client.beta.assistants.create(
        name="ORMIT Report Assessor",
        instructions="You are a senior trainee assessor at a Belgian company ORMIT Talent. Your task is to extract and provide assessment data for a trainee based on notes from assessors who met the trainee and a personality and cognitive capacity test. You also have a context/task elaboration file and a file specifying the tone of voice for this.",
        model="gpt-4o-mini",
        tools=[{"type": "file_search"}],
    )
    
    # Ensure that the assistant has been created and you can access its ID
    if not hasattr(assistant, 'id'):
        raise Exception("Assistant creation failed or ID not found.")
    
    global_signals.update_message.emit("Succesfully connected, uploading files...")

    # Create the vector store
    vector_store = client.beta.vector_stores.create(name="Assessment_Data", expires_after={
        "anchor": "last_active_at",
        "days": 1
    })
    
    # Prepare files for upload
    file_streams = [open(path, "rb") for path in lst_files]
    
    # Upload files to the vector store
    file_batch = client.beta.vector_stores.file_batches.upload_and_poll(
        vector_store_id=vector_store.id, files=file_streams
    )
    
    # Close file streams after upload
    for file in file_streams:
        file.close()
    
    print(file_batch.file_counts)
    
    # (Re)connect the assistant to the vector store(s)
    assistant = client.beta.assistants.update(
        assistant_id=assistant.id,
        tool_resources={"file_search": {"vector_store_ids": [vector_store.id]}},
    )
    
    global_signals.update_message.emit("Files uploaded, starting prompts...")
    
    # Now you should have access to the updated assistant's ID
    assistID = assistant.id
    print(f"Assistant ID: {assistID}")
    
    # Collect prompts
    lst_prompts = get_custom_key_list(prompts)
    print(lst_prompts)
    
    for promno, prom in enumerate(lst_prompts, start=1):
        print(prom)
        global_signals.update_message.emit(f"Processing prompt {promno}/{len(lst_prompts)}, please wait...")

        # Modify the sleep command for these specific prompts
        if prom in ['prompt4_cogcap', 'prompt6a_conqual', 'prompt6b_conimprov']:
            time.sleep(90)  # Increase from 61 to 90 seconds for these three
               
        # Create a new thread for each prompt
        empty_thread = client.beta.threads.create()
        
        # Create a message in the new thread
        client.beta.threads.messages.create(
            empty_thread.id,
            role="user",
            content=prompts[prom],
        )
        
        # Run the assistant
        run = client.beta.threads.runs.create(
            thread_id=empty_thread.id,
            assistant_id=assistID,  # Use assistID here
        )
        
        start_wait_time = time.time()

        while run.status != 'completed':
            time.sleep(2)  # Avoid overloading the server
            run = client.beta.threads.runs.retrieve(thread_id=empty_thread.id, run_id=run.id)
            
            # Check if wait time exceeds the max limit
            if time.time() - start_wait_time > max_wait_time:
                print(f"Timeout for {prom}")
                output = ''
                break
            # Only proceed if the run is completed
            if run.status == 'completed':
                output = client.beta.threads.messages.list(thread_id=empty_thread.id, run_id=run.id)
                messages = list(output)
                message_content = messages[0].content[0].text
                output = message_content.value
                            
        output_label = prom

        results[output_label] = output
        print(f"{output_label}: {output}")
        
        appl_name = data["Applicant Name"]
        filename_with_timestamp = f"{appl_name}_{formatted_time}.json"  # Filename with timestamp
        with open(filename_with_timestamp, 'w') as json_file:
            json.dump(results, json_file, indent=4)  # 'indent=4' for pretty printing
            
    # Clean vector store
    file_ids = {file.id for file in client.files.list()}
    print(file_ids)
    
    # Then attempt deletion only for files that are still in the list
    for file_id in file_ids:
        client.files.delete(file_id)
        print(f"Deleted file {file_id}")
    client.beta.assistants.delete(assistant.id)
    
    global_signals.update_message.emit(f"Prompting finished, generating report...")
    return filename_with_timestamp