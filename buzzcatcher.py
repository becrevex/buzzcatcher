from docx import Document
from docx.enum.text import WD_COLOR_INDEX
import re

phrases = [
    "At the end of the day", "With that being said", "It goes without saying", "In a nutshell",
    "Needless to say", "When it comes to", "A significant number of", "It’s worth mentioning",
    "Last but not least", "Cutting‑edge", "Leveraging", "Moving forward", "Going forward",
    "On the other hand", "Notwithstanding", "Takeaway", "As a matter of fact", "In the realm of",
    "Seamless integration", "Robust framework", "Holistic approach", "Paradigm shift", "Synergy",
    "Scale-up", "Optimize", "Game‑changer", "Unleash", "Uncover", "In a world", "In a sea of",
    "Digital landscape", "Elevate", "Embark", "Delve", "Game Changer", "In the midst", "In addition"
]

normalized_phrases = [p.lower() for p in phrases]

def highlight_phrases(doc):
    for para in doc.paragraphs:
        for phrase in phrases:
            # keeping it non-case-sensitive with re
            matches = re.finditer(re.escape(phrase), para.text, re.IGNORECASE)
            for match in matches:
                start, end = match.start(), match.end()
                highlight_text_in_run(para, start, end)
    return doc

def highlight_text_in_run(paragraph, start_index, end_index):
    current_index = 0
    for run in paragraph.runs:
        run_length = len(run.text)
        if current_index + run_length < start_index:
            current_index += run_length
            continue
        if current_index <= start_index < current_index + run_length:
            match_offset = start_index - current_index
            end_offset = min(end_index - current_index, run_length)
            run.font.highlight_color = WD_COLOR_INDEX.YELLOW
        current_index += run_length

doc = Document("your_document.docx")
highlighted_doc = highlight_phrases(doc)
highlighted_doc.save("highlighted_output.docx")
