from Code import slide
import Code.imp_func as ip
import pandas as pd

class RiskDeckSlide(slide.AbstractRiskDeckSlide):

    file_name= "Koalafi Lease Data.xlsx"
    slide_title = 'Final State'
    
    def build_slide(self):
        #getting current slide and data
        chart_objects = [x for x in self.current_slide.shapes if x.has_chart]
        df = self.data

        #transforming the data
        df['APP_CREATE_DATE'] = pd.to_datetime(df['APP_CREATE_DATE'])
        df['YEAR'] = df['APP_CREATE_DATE'].dt.year
        df['MONTH'] = df['APP_CREATE_DATE'].dt.month

        #chart 1
        chart1_df = df.groupby(pd.Grouper(key = 'APP_CREATE_DATE',freq='1M'))[['FUNDED_COUNT','CHARGEOFF','EARLY_BUY_OUT']].sum().reset_index()
        chart1_df['EBO Rate'] = chart1_df['EARLY_BUY_OUT']/chart1_df['FUNDED_COUNT']
        chart1_df['CO Rate'] = chart1_df['CHARGEOFF']/chart1_df['FUNDED_COUNT']
        chart1_df['Paid in Full Rate'] = 1 - chart1_df['EBO Rate'] - chart1_df['CO Rate']
        chart1_df['APP_CREATE_DATE'] = chart1_df['APP_CREATE_DATE'].dt.strftime('%m-%Y')
        chart_cols = ['EBO Rate','CO Rate','Paid in Full Rate']
        ip.update_chart(chart_objects[0],chart1_df, 'APP_CREATE_DATE',chart_cols)



