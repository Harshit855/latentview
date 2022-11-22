import pptx
from datetime import date

def Title_date(slide):
    
    today = date.today()
    today = today.strftime('%b %d, %Y')
    
    title_date = [shape for shape in slide.shapes if shape.has_text_frame and shape.text == '25th July 2022']
    title_date[0].text = today

    return