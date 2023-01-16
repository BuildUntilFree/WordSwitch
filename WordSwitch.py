import docx
import os

# The keyword you want to replace
old_keyword = 'old'
# The new keyword you want to use
new_keyword = 'new'

# Specify the directory containing the .docx files
directory = r'C:\FileLocation\WordSwitch\From'

# Iterate through all files in the directory
for filename in os.listdir(directory):
    # Only open .docx files
    if filename.endswith('.docx'):
        # Open the .docx file
        doc = docx.Document(os.path.join(directory, filename))

        # Iterate through all paragraphs in the document
        for p in doc.paragraphs:
            # Replace the keyword in the text
            p.text = p.text.replace(old_keyword, new_keyword)

        # Save the changes to the .docx file
        doc.save(os.path.join(r'C:\FileLocation\WordSwitch\to', filename))

print("WordSwitched!")
