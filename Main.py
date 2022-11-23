import pip
pip.main(['install','python-pptx'])
pip.main(['install','pandas'])

import importlib
from Code.Title_slide import Title_date
from datetime import datetime
import pptx
import os

slides_file = os.listdir('Code/Slides/')
Slides = []
for i in slides_file:
    if i != '__pycache__':
        slide_module = importlib.import_module(f'Code.Slides.{i[:-3]}')
        Slides.append(slide_module.RiskDeckSlide)

def make_slides(prez):
    Title_date(slide=prez.slides[0])
    for Slide in Slides:
        try:
            Slide(pptx_prez=prez).build_slide()
        except Exception as e:
            print(f'{Slide} Failed')

    
def main(ppt='Template Risk Deck - Overall.pptx'):
    prez = pptx.Presentation(ppt)
    date_today = datetime.today().strftime('%Y-%m-%d')

    make_slides(prez)
    prez.save(f"{date_today} Risk Deck.pptx")

if __name__ == '__main__':
    main()
