
import pip
pip.main(['install','seaborn'])
pip.main(['install','matplotlib'])
from Code import slide
import Code.imp_func as ip
import pandas as pd
import seaborn as sns
import matplotlib.pyplot as plt
from pptx.util import Inches

class RiskDeckSlide(slide.AbstractRiskDeckSlide):

    file_name= "Koalafi Lease Data.xlsx"
    slide_title = 'Total vs Funded Applications'
    
    def build_slide(self):
        #getting current slide and data
        df = self.data

        #Creating the graph and saving it
        df['APP_CREATE_DATE'] = pd.to_datetime(df['APP_CREATE_DATE'])
        chart_df = df.groupby(["DEALER_NAME"])[['TOTAL_APPS','FUNDED_COUNT']].sum().reset_index()
        sns.lmplot(data=chart_df,x='TOTAL_APPS',y='FUNDED_COUNT',order=2,ci=90)
        plt.savefig('Total vs Funded Applications.png')
        
        # add images to slide and save
        self.current_slide.shapes.add_picture('Total vs Funded Applications.png', Inches(2.75), Inches(1.05), height=Inches(6.25), width=Inches(7.75))
