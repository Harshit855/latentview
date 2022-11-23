
from Code import slide
import Code.imp_func as ip
import pandas as pd

class RiskDeckSlide(slide.AbstractRiskDeckSlide):

    file_name= "Koalafi Lease Data.xlsx"
    slide_title = 'Pre-Paid Fraud Rate'
    
    def build_slide(self):
        #getting current slide and data
        chart_objects = [x for x in self.current_slide.shapes if x.has_chart]
        df = self.data

        #transforming the data
        df['APP_CREATE_DATE'] = pd.to_datetime(df['APP_CREATE_DATE'])
        df['YEAR'] = df['APP_CREATE_DATE'].dt.year
        df['MONTH'] = df['APP_CREATE_DATE'].dt.month
        data = df.groupby(['MONTH',"YEAR"])[['PREPAID_CARD_FRAUD','FUNDED_COUNT']].sum().reset_index()
        data['Prepaid_fruad'] = data['PREPAID_CARD_FRAUD']/data['FUNDED_COUNT']

        #updating chart 1
        chart1_df = pd.pivot_table(data = data,index='MONTH',columns='YEAR',values='Prepaid_fruad').reset_index()
        chart1_df = ip.transform_month_number(chart1_df)
        chart_cols = chart1_df.columns.values.tolist()[1:]
        ip.update_chart(chart_objects[0],chart1_df,"Application Month",chart_cols)
