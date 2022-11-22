from Code import slide
import Code.imp_func as ip
import pandas as pd

class RiskDeckSlide(slide.AbstractRiskDeckSlide):

    file_name= "Koalafi Lease Data.xlsx"
    slide_title = 'Delinquency Rate'
    
    def build_slide(self):
        #getting current slide and data
        chart_objects = [x for x in self.current_slide.shapes if x.has_chart]
        df = self.data

        #transforming the data
        df['APP_CREATE_DATE'] = pd.to_datetime(df['APP_CREATE_DATE'])
        data = df.groupby(pd.Grouper(key='APP_CREATE_DATE',freq='1M'))[['DQ_1PLUS','DQ_30PLUS','DQ_60PLUS','DQ_120PLUS','FUNDED_COUNT']].sum().reset_index()
        data['DQ_1+'] = data['DQ_1PLUS']/data['FUNDED_COUNT']
        data['DQ_30+'] = data['DQ_30PLUS']/data['FUNDED_COUNT']
        data['DQ_60+'] = data['DQ_60PLUS']/data['FUNDED_COUNT']
        data['DQ_120+'] = data['DQ_120PLUS']/data['FUNDED_COUNT']

        #updating chart 1
        chart1_df = data.copy()
        chart1_df['APP_CREATE_DATE'] = chart1_df['APP_CREATE_DATE'].dt.strftime('%m-%Y')
        chart1_cols = ['DQ_1+','DQ_30+','DQ_60+','DQ_120+']
        ip.update_chart(chart_objects[0],chart1_df,"APP_CREATE_DATE",chart1_cols)
