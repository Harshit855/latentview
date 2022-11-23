
from Code import slide
import Code.imp_func as ip
import pandas as pd

class RiskDeckSlide(slide.AbstractRiskDeckSlide):

    file_name= "Koalafi Lease Data.xlsx"
    slide_title = 'Executive Overview'
    
    def build_slide(self):
        #getting current slide and data
        table_objects = [x for x in self.current_slide.shapes if x.has_table]
        df = self.data

        #transforming the data
        df['APP_CREATE_DATE'] = pd.to_datetime(df['APP_CREATE_DATE'])
        df['YEAR'] = df['APP_CREATE_DATE'].dt.year
        df['MONTH'] = df['APP_CREATE_DATE'].dt.month

        #Table 1
        data = df.groupby(pd.Grouper(key = 'APP_CREATE_DATE',freq='1M'))[['TOTAL_APPS','HEADLINE_APP','PRE_APPROVAL','NET_APPROVAL','FUNDED_COUNT']].sum().reset_index()
        data['Initial Approval Rate'] = data['HEADLINE_APP']/data['TOTAL_APPS']
        data['Pre Approval Rate'] = data['PRE_APPROVAL']/data['TOTAL_APPS']
        data['Net Approval Rate'] = data['NET_APPROVAL']/data['TOTAL_APPS']
        table1_df = data[['TOTAL_APPS','Initial Approval Rate','Pre Approval Rate','Net Approval Rate','FUNDED_COUNT']].iloc[-2:]
        for col in ['Initial Approval Rate','Pre Approval Rate','Net Approval Rate']:
            table1_df[col] = round(table1_df[col]*100,1)
        for col in table1_df.columns:
            table1_df[col] = table1_df[col].astype('str')
        table1_df = table1_df.T
        ip.write_table(table_objects[0].table,table1_df,1,1)

        #Table 2
        data = df.groupby(pd.Grouper(key = 'APP_CREATE_DATE',freq='1M'))[['FUNDED_COUNT','MISSED_FIRST_PAYMENT','MISSED_SECOND_PAYMENT','STATEMENT1_1PLUS','STATEMENT2_1PLUS','STATEMENT3_30PLUS']].sum().reset_index()
        data['MFP Rate'] = data['MISSED_FIRST_PAYMENT']/data['FUNDED_COUNT']
        data['MSP Rate'] = data['MISSED_SECOND_PAYMENT']/data['FUNDED_COUNT']
        data['Stmt1_1'] = data['STATEMENT1_1PLUS']/data['FUNDED_COUNT']
        data['Stmt2_1'] = data['STATEMENT2_1PLUS']/data['FUNDED_COUNT']
        data['Stmt3_30'] = data['STATEMENT3_30PLUS']/data['FUNDED_COUNT']
        table2_df = data[['MFP Rate','MSP Rate','Stmt1_1','Stmt2_1','Stmt3_30']].iloc[-2:]
        for col in table2_df.columns:
            table2_df[col] = round(table2_df[col]*100,1).astype('str')
        table2_df = table2_df.T
        ip.write_table(table_objects[1].table,table2_df,1,1)





