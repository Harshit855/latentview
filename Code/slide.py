
import pip
pip.main(['install','pandas'])
pip.main(['install','attr'])

from abc import ABC, abstractmethod
from attr import attr
import pandas as pd
import os

class AbstractRiskDeckSlide(ABC):

    def __init__(self, pptx_prez):
        self.prez = pptx_prez
        
        for slide_name in self.prez.slides:
            if slide_name.shapes.title.text == self.slide_title:
                self.current_slide = slide_name
        self.get_data()
    
    def get_data(self):
        if self.file_name is not None:
            file = os.path.join("Code","Data",self.file_name)
            self.data = pd.read_excel(file, sheet_name='Raw Data',index_col=None)

    def build_slide(self):
        pass
