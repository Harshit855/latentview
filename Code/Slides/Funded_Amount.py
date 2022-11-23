
from Code import slide
import Code.imp_func as ip
import pandas as pd

class RiskDeckSlide(slide.AbstractRiskDeckSlide):

    file_name= "Koalafi Lease Data.xlsx"
    slide_title = 'Funded Amount'
    
    def build_slide(self):
        #getting current slide and data
        chart_objects = [x for x in self.current_slide.shapes if x.has_chart]
        df = self.data

        #transforming the data
        df['APP_CREATE_DATE'] = pd.to_datetime(df['APP_CREATE_DATE'])
        df['YEAR'] = df['APP_CREATE_DATE'].dt.year
        df['MONTH'] = df['APP_CREATE_DATE'].dt.month

        #chart 1
        chart1_df = df.groupby(pd.Grouper(key = 'APP_CREATE_DATE',freq='1M'))[['FUNDED_COUNT','LOANED_AMOUNT']].mean().reset_index()
        chart1_df['APP_CREATE_DATE'] = chart1_df['APP_CREATE_DATE'].dt.strftime('%m-%Y')
        chart_cols = ['FUNDED_COUNT','LOANED_AMOUNT']
        ip.update_chart(chart_objects[0],chart1_df, 'APP_CREATE_DATE',chart_cols)



