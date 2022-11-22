from Code import slide
import Code.imp_func as ip
import pandas as pd

class RiskDeckSlide(slide.AbstractRiskDeckSlide):

    file_name= "Koalafi Lease Data.xlsx"
    slide_title = 'Banking Soft & Hard Decline Rate'
    
    def build_slide(self):
        #getting current slide and data
        chart_objects = [x for x in self.current_slide.shapes if x.has_chart]
        df = self.data

        #transforming the data
        df['APP_CREATE_DATE'] = pd.to_datetime(df['APP_CREATE_DATE'])
        df['YEAR'] = df['APP_CREATE_DATE'].dt.year
        df['MONTH'] = df['APP_CREATE_DATE'].dt.month
        data = df.groupby(['MONTH',"YEAR"])[['BAV_FRAUD_SOFT','BAV_FRAUD_HARD','FUNDED_COUNT']].sum().reset_index()
        data['BAV_Soft_Decline'] = data['BAV_FRAUD_SOFT']/data['FUNDED_COUNT']
        data['BAV_Hard_Decline'] = data['BAV_FRAUD_HARD']/data['FUNDED_COUNT']

        #updating chart 1
        chart1_df = pd.pivot_table(data = data,index='MONTH',columns='YEAR',values='BAV_Soft_Decline').reset_index()
        chart1_df = ip.transform_month_number(chart1_df)
        chart1_cols = chart1_df.columns.values.tolist()[1:]
        ip.update_chart(chart_objects[0],chart1_df,"Application Month",chart1_cols)

        #updating chart 2
        chart2_df = pd.pivot_table(data = data,index='MONTH',columns='YEAR',values='BAV_Hard_Decline').reset_index()
        chart2_df = ip.transform_month_number(chart2_df)
        chart2_cols = chart2_df.columns.values.tolist()[1:]
        ip.update_chart(chart_objects[1],chart2_df,"Application Month",chart2_cols)