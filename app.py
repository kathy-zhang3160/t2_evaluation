import streamlit as st
import pandas 
import numpy
import io
import xlsxwriter

def compliance_score(df, with_issue, without_issue):
    df['terminate'] = 'N'

    df.loc[df['抓货w/issue'] > 1, 'terminate'] = 'terminate'
    df.loc[df['抓货wo/issue'] > 2, 'terminate'] = 'PIP'
    df['cplc_score'] = 1 - (with_issue * df['抓货w/issue'] + without_issue * df['抓货wo/issue']*0.5)
    return df['cplc_score'], df['terminate']

def uploaded_weight_cal(esc_cplc, df, df_weight):
    # df_weight= pandas.read_excel('weight_dic.xlsx')
    weight_L3 = df_weight.query('level == "level3" and mix == "Y"')
    L3_dict = weight_L3.set_index('index')['weight'].to_dict()

    erp_coef = df_weight.query('level == "level3" and mix == "coef"')
    erp_coef = erp_coef.set_index('index')['weight'].to_dict()
    for col_name, weight in L3_dict.items():
        df[f'{col_name}_pct'] = (df[col_name] / df[col_name].sum())
        df[f'{col_name}_mix'] = (df[col_name] / df[col_name].sum())*weight
        
    with_issue = df_weight[df_weight['index'] == "抓货w/issue"]['weight'].values[0]
    without_issue = df_weight[df_weight['index'] == "抓货wo/issue"]['weight'].values[0]
    # 二级指标
    df['sales_team_L2'] = df['AE#_mix'] + df['FTE#_mix']
    
    col_select = [col for col in df.columns if col.startswith('FTE# Tier') and col.endswith('mix')]
    df['region_cover_L2'] = df[col_select].sum(axis=1)
    
    df['top_acct_L2'] = df['top_account#_p4q_mix'] + df['top_account#_cq_mix']
    
    df['active_acct_L2'] = df['active_account#_mix'] + df['account_order#_mix']
    
    df['esc_store_L2'] = df['ESC#_p4q_mix'] + df['esc_store#_cq_mix']
    
    df['ec_store_L2'] = df['ec_store#_p4q_mix'] + df['ec_store#_cq_mix']
    
    # ERP connect score
    df['erp_score_L2'] = erp_coef["ERP_base"]
    df.loc[df['ERP_conn'] == '直连互道', 'erp_score_L2'] = erp_coef["ERP_direct"]
    df.loc[df['ERP_conn'] == '云开中转', 'erp_score_L2'] = erp_coef['ERP_indirect']
    df['erp_score_L2'] = df['erp_score_L2']/ df['erp_score_L2'].sum()
    
    col_select = [col for col in df.columns if col.startswith('rev') and col.endswith('mix')]
    df['sales_revenue_L2'] = df[col_select].sum(axis = 1)
    
    ## Compliance Score
    df['sales_compliance_L2'],df['t2_terminate'] = compliance_score(df, with_issue, without_issue)
    
    esc_cplc['esc_compliance_L2'], esc_cplc['esc_terminate'] = compliance_score(esc_cplc, with_issue, without_issue)
    
    esc_cplc = esc_cplc.groupby('reseller').agg({'esc_terminate':'max',
                                      'esc_compliance_L2':'mean'}).reset_index()
    df = df.merge(esc_cplc, on = 'reseller', how = 'left')
    
    # special deal for esc compliance L2, 因为全部应该为1，扣分制，空值应为1
    df['esc_compliance_L2'] = df['esc_compliance_L2'].fillna(1)
    
    df['sales_compliance_L2'] = df['sales_compliance_L2']/df['sales_compliance_L2'].sum()
    
    df['esc_compliance_L2'] = df['esc_compliance_L2']/df['esc_compliance_L2'].sum()

    weight_L2 = df_weight.query('level == "level2"')
    L2_dict = weight_L2.set_index('index')['weight'].to_dict()

    for col_name, weight in L2_dict.items():
        df[f'{col_name}_wt'] = (df[col_name])*weight
    
   # 一级指标

    df['loyalty_L1'] = df['service_period']/df['service_period'].sum()
    
    df['team_L1'] = df['sales_team_L2_wt'] + df['region_cover_L2_wt']
    
    df['account_L1'] = df['top_acct_L2_wt'] + df['active_acct_L2_wt']
    
    df['program_L1'] = df['esc_store_L2_wt'] + df['ec_store_L2_wt'] + df['erp_score_L2_wt']
    
    df['sales_L1'] = df['sales_revenue_L2_wt']
    
    df['compliance_L1'] = df['sales_compliance_L2_wt'] + df['esc_compliance_L2_wt']
    
    weight_L1 = df_weight.query('level == "level1"')
    L1_dict = weight_L1.set_index('index')['weight'].to_dict()

    for col_name, weight in L1_dict.items():
        df[f'{col_name}_wt'] = (df[col_name])*weight
    df['T2_evaluation'] = df[[col for col in df.columns if col.endswith('L1_wt')]].sum(axis = 1)
    key_index = ['status', 'tier', 'disti_t1', 'reseller', 'onboard_FYQ',
           'terminate_FYQ', 'T2_evaluation', 't2_terminate']
    df = df[[col for col in df.columns if (col in (key_index) or col.endswith('pct') or col.endswith('L2') or col.endswith('L1') )]]
    df = df.sort_values(by = 'T2_evaluation', ascending = False)
    return df


st.title('Excel Processor')

uploaded_file = st.file_uploader("Upload an Excel file", type="xlsx")

if uploaded_file is not None:
    df_weight = pandas.read_excel(uploaded_file)
    url = 't2_evaluation_raw_data.xlsx'
    esc_cplc = pandas.read_excel('esc_cplc_output.xlsx', engine = 'openpyxl')
    df = pandas.read_excel(url, engine='openpyxl')
    df = uploaded_weight_cal(esc_cplc, df, df_weight)
    if st.checkbox('Preview dataframe'):
        st.dataframe(df)
    st.subheader('Result Preview:')
    st.write(df.head())
    buffer = io.BytesIO()
    with pandas.ExcelWriter(buffer, engine='xlsxwriter') as writer:
        df.to_excel(writer, sheet_name='Sheet1', index = False)

        # Close the Pandas Excel writer and output the Excel file to the buffer
        writer.close()

        st.download_button(
            label="Download Excel worksheets",
            data=buffer,
            file_name="t2_evaluation.xlsx",
            mime="application/vnd.ms-excel"
        )
    # download to excel 
    
    # st.write("Processed Data:")
    # st.dataframe(df)

    # # 下载链接
    # st.download_button(
    #     label="Download Processed Excel",
    #     data=df.to_excel(index=False, engine='openpyxl'),
    #     file_name='output.xlsx',
    #     mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    # )
