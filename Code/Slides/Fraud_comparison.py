
import pip
pip.main(['install','pandas'])
from Code import slide
import Code.imp_func as ip
import pandas as pd

class RiskDeckSlide(slide.AbstractRiskDeckSlide):

    file_name= "Koalafi Lease Data.xlsx"
    slide_title = 'Fraud Comparison'
    
    def build_slide(self):
        #getting current slide and data
        chart_objects = [x for x in self.current_slide.shapes if x.has_chart]
        df = self.data

        #transforming the data
        df['APP_CREATE_DATE'] = pd.to_datetime(df['APP_CREATE_DATE'])
        df['YEAR'] = df['APP_CREATE_DATE'].dt.year
        df['MONTH'] = df['APP_CREATE_DATE'].dt.month

        #chart 1
        chart1_df = df.groupby('VERTICAL')[['IDV_FRAUD','BAV_FRAUD','PREPAID_CARD_FRAUD','TOTAL_APPS']].sum().reset_index()
        chart_cols = chart1_df.columns.values.tolist()[1:]
        ip.update_chart(chart_objects[0],chart1_df, 'VERTICAL',chart_cols)



