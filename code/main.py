from harcode import *

def call_ta_report(period='12-2025'):
    def get_monthly_return():
        def get_monthly_vni_return():
            df_vni = pd.read_excel(raw_filepath, sheet_name=vni_sheetname)
            df_vni['month'] = df_vni['Date'].dt.to_period('M')
            
            last_month_df = df_vni.groupby(by='month', as_index=False).last()
            last_month_df['return'] = last_month_df['VNINDEX'].pct_change()
            
            df = last_month_df[last_month_df['return'].notna()]
            return df
        
        def get_monthly_port_return():
            df_port = pd.read_excel(raw_filepath, sheet_name=port_sheetname)
            df_port = df_port[df_port['Close_Date'].notna()]
            df_port['month'] = df_port['Close_Date'].dt.to_period('M')
            df_port['return'] = df_port['Close_Price']/df_port['Open_Price'] - 1
            
            df = df_port.groupby(by='month', as_index=False)['return'].mean()
            return df
        
        df_vni = get_monthly_vni_return()
        df_port = get_monthly_port_return()
        df = pd.merge(df_vni, df_port, how='left', on='month', suffixes=['_vni','_port'])
        
        df = pd.concat([pd.DataFrame([{'month': f'{int(period[-4:]) - 1}-12',
                                       'return_vni': 0,
                                       'return_port': 0,
                                       'index_vni': 100,
                                       'index_port': 100}]), df], ignore_index=True)

        for i in range(1,len(df)):
            df.loc[i, 'index_vni'] = (df.loc[i, 'return_vni'] + 1) * 100
            df.loc[i, 'index_port'] = (df.loc[i, 'return_port'] + 1) * 100
        
        df.loc[df['return_vni'].isna(), 'index_vni'] = 100
        df.loc[df['return_port'].isna(), 'index_port'] = 100

        df['month'] = df['month'].astype(str)
        return df[['month', 'return_vni', 'return_port', 'index_vni', 'index_port']]
    
    def get_mtd_return():
        df_vni = pd.read_excel(raw_filepath, sheet_name=vni_sheetname)
        df_port = pd.read_excel(raw_filepath, sheet_name=port_sheetname)
        def get_mtd_vni_index():
            previous_period = f'{int(period[:2]) - 1}-{period[-4:]}' if int(period[:2]) > 1 else f'12-{int(period[-4:]) - 1}'
            df_vni = pd.read_excel(raw_filepath, sheet_name=vni_sheetname)
            df_vni['change'] = df_vni['VNINDEX'].pct_change()
            df_vni['month'] = df_vni['Date'].dt.to_period('M')
            df_vni_arr = [df_vni[df_vni['month'] == previous_period].tail(1),
                        df_vni[df_vni['month'] == period]]
            df_vni = pd.concat(df_vni_arr, ignore_index=True)
            df_vni.loc[0,'index_vni'] = 100
            for i in range(1,len(df_vni)):
                df_vni.loc[i,'index_vni'] = (1 + df_vni.loc[i,'change'])*df_vni.loc[i-1,'index_vni']
            return df_vni[['Date', 'index_vni']]
        
        def get_mtd_port_index():
            df_port = pd.read_excel(raw_filepath, sheet_name=port_sheetname)
            df_port = df_port[df_port['Close_Date'].notna()]
            df_port['month'] = df_port['Close_Date'].dt.to_period('M')
            df_port = df_port[df_port['month'] == period]

            df_port['return'] = df_port['Close_Price']/df_port['Open_Price'] - 1
            df_port = df_port.groupby(by='Close_Date')['return'].mean().reset_index()
            
            
            for i in range(len(df_port)):
                if i == 0:
                    df_port.loc[i, 'index_port'] = (1 + df_port.loc[i, 'return'])*100
                else:
                    df_port.loc[i, 'index_port'] = (1 + df_port.loc[i, 'return'])*df_port.loc[i-1, 'index_port']

            df_port['Date'] = df_port['Close_Date']
            return df_port[['Date', 'index_port']]

        df_vni = get_mtd_vni_index()
        df_port = get_mtd_port_index()
        df = pd.merge(df_vni, df_port, how='left', on='Date')
        df['index_port'] = df['index_port'].ffill()
        df['index_port'] = df['index_port'].fillna(100)
        return df

    def get_mtd_statistics():
        df_port = pd.read_excel(raw_filepath, sheet_name=port_sheetname)
        df_port = df_port[df_port['Close_Date'].notna()]
        df_port['month'] = df_port['Close_Date'].dt.to_period('M')
        df_port = df_port[df_port['month'] == period]

        df_port['return'] = df_port['Close_Price']/df_port['Open_Price'] - 1  
        df_port['month'] = df_port['month'].astype(str)
        df_port = df_port.sort_values(by='return',ascending=False)     
        return df_port

    wb = xw.Book(output_wb)

    def data_to_excel(df, sheet_name):
        ws = wb.sheets[sheet_name]

        ws.used_range.clear_contents()
        ws['A1'].options(index=False).value = df

        wb.save()

    data_to_excel(
        df = get_mtd_return(),
        sheet_name="mtd"
    )
    data_to_excel(
        df = get_mtd_statistics(),
        sheet_name="monthly_sta"
    )
    data_to_excel(
        df = get_monthly_return(),
        sheet_name="monthly"
    )

call_ta_report(period='12-2025')