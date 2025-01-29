import fitz
import re
import os
import shutil

class Redactor:
    # static methods work independent of class object
    @staticmethod
    def get_sensitive_data(lines, target_names):
        """ Function to get sensitive data lines containing specified keywords and other sensitive information """
        # Create a regex pattern for names (case insensitive)
        NAME_REG = r'\b(' + '|'.join(re.escape(name) for name in target_names) + r')\b'
        # Regex pattern for email addresses
        EMAIL_REG = r'[\w\.-]+@[\w\.-]+'
        # Regex pattern for phone numbers starting with a country code
        PHONE_REG = r'\+\d{1,3}\s*\d{1,3}(\s*\d{2,3}){2,4}'  # Matches formats like +32 496 61 73 89 or +31 6 1234 5678
        
        # Keywords to trigger line redaction
        keywords = ["gender", "address", "phone", "e-mail", "date of birth", "links", "socials"]

        previous_line = None  # Variable to store the previous line

        for line in lines:
            # Check if the previous line starts with 'Address'
            for keyword in keywords: 
                if previous_line and previous_line.lower().startswith(keyword):
                    yield line  # Print the current line if the previous line started with 'address'

            # Update the previous line
            previous_line = line
            
            # matching names
            for match in re.finditer(NAME_REG, line, re.IGNORECASE):
                yield match.group(0)  # yield matched name
            
            # matching email addresses
            for match in re.finditer(EMAIL_REG, line):
                yield match.group(0)  # yield matched email address
            
            # matching phone numbers
            for match in re.finditer(PHONE_REG, line):
                yield match.group(0)  # yield matched phone number

    # constructor
    def __init__(self, path, target_names, new_filename, profile_pic):
        self.path = path
        self.target_names = target_names
        self.profile_pic = profile_pic

    def redaction(self, new_filename):
        """ main redactor code """

        # opening the pdf
        doc = fitz.open(self.path)
        # delete metadata
        doc.set_metadata({})  # Removes all metadata

        profile_pic_rect = fitz.Rect(10, 10, 65, 75)  # Adjust as per the exact position and size
        
        # iterating through pages
        for i, page in enumerate(doc):
            if self.profile_pic and i == 0:
                print(f"Redacting profile picture on page {i + 1}")  # Debugging message
                page.add_redact_annot(profile_pic_rect, fill=(0, 0, 0))  # Redact the profile picture area
            
            # Get text from the page
            text = page.get_text("text").split('\n')  # Updated to get_text()
            # getting the rect boxes which consists the matching names, email, and phone regex
            sensitive = self.get_sensitive_data(text, self.target_names)
            for data in sensitive:
                print(f"Sensitive data found: {data}")  # Debugging line
                areas = page.search_for(data)
                print(f"Redaction areas for '{data}': {areas}")  # Debugging line

                # drawing outline over sensitive data
                [page.add_redact_annot(area, fill=(0, 0, 0)) for area in areas]  # Updated to add_redact_annot()

            # applying the redaction
            page.apply_redactions()

        # saving it to a new pdf
        output_path = os.path.join('temp', f"{new_filename}.pdf")  # Update to save in temp folder
        doc.save(output_path)
        
        print(f"Successfully redacted {self.path} to {output_path}")

def create_temp_folder():
    # Define the temporary folder path
    temp_folder = 'temp'
    # Create the temp directory if it doesn't exist
    if not os.path.exists(temp_folder):
        os.makedirs(temp_folder)
            
# driver code for testing
def redact_folder(data, profile_pic=False):
    create_temp_folder() #Create Temp
    full_name_split = data["Applicant Name"].split()  # Split the name by spaces
    target_names = [full_name_split[0], full_name_split[-1]] #In case of name with more than 2 parts, keep first and last for redaction
    print(target_names)
    for new_filename, filename in data["Files"].items():
        path = os.path.join('temp', filename)  # Construct the path to the file
        if new_filename == "Assessment Notes":
            profile_pic=True
        else:
            profile_pic=False
        redactor = Redactor(path, target_names, new_filename, profile_pic)  # Create a Redactor instance
        redactor.redaction(new_filename)  # Perform redaction