
from Code import slide
import Code.imp_func as ip
import pandas as pd

class RiskDeckSlide(slide.AbstractRiskDeckSlide):

    file_name= "Koalafi Lease Data.xlsx"
    slide_title = 'Missed Payment Metric'
    
    def build_slide(self):
        #getting current slide and data
        chart_objects = [x for x in self.current_slide.shapes if x.has_chart]
        df = self.data

        #transforming the data
        df['APP_CREATE_DATE'] = pd.to_datetime(df['APP_CREATE_DATE'])
        df['YEAR'] = df['APP_CREATE_DATE'].dt.year
        df['MONTH'] = df['APP_CREATE_DATE'].dt.month

        #chart 1
        chart1_df = df.groupby('PAYMENT_FREQUENCY')[['MISSED_FIRST_PAYMENT','MISSED_SECOND_PAYMENT']].sum().reset_index()
        chart1_cols = chart1_df.columns.values.tolist()[1:]
        ip.update_chart(chart_objects[0],chart1_df, 'PAYMENT_FREQUENCY',chart1_cols)



