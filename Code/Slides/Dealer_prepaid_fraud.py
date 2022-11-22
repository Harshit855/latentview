
import pip
pip.main(['install','pandas'])
from Code import slide
import Code.imp_func as ip
import pandas as pd

class RiskDeckSlide(slide.AbstractRiskDeckSlide):

    file_name= "Koalafi Lease Data.xlsx"
    slide_title = 'Top 10 Dealers By Pre-Paid Fraud'
    
    def build_slide(self):
        #getting current slide and data
        chart_objects = [x for x in self.current_slide.shapes if x.has_chart]
        df = self.data

        #transforming the data
        df['APP_CREATE_DATE'] = pd.to_datetime(df['APP_CREATE_DATE'])

        #updating chart 1
        chart1_df = df.groupby('DEALER_NAME')['PREPAID_CARD_FRAUD'].sum().reset_index()
        chart1_df = chart1_df.sort_values(by=['PREPAID_CARD_FRAUD'],ascending=False).iloc[:10].reset_index(drop=True)
        chart1_df = chart1_df.iloc[::-1] #Revesing the table
        chart1_cols = ['PREPAID_CARD_FRAUD']
        ip.update_chart(chart_objects[0],chart1_df,"DEALER_NAME",chart1_cols)
