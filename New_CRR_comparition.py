import pandas as pd
import numpy as np
import glob

col_name = ['Date','Asset_id','contracts','underlying','length']

# Get CRR_1 and CRR_2 files from particular location and remove bracket from path using Indiexing
file_1 = glob.glob(r'C:\Users\dls\Desktop/python/CODE/New folder/*_1')  
file_1 = file_1[0]                                                    
file_2 = glob.glob(r'C:\Users\dls\Desktop/python/CODE/New folder/*_2') 
file_2 = file_2[0]                                                     

# Read both CRR files using pandas function
crr_1= pd.read_csv(file_1,sep='|',header=None,names=col_name,usecols=['Asset_id','contracts','underlying'])
crr_2 = pd.read_csv(file_2,sep='|',header=None,names=col_name,usecols=['Asset_id','contracts','underlying'])

# Store total count of CRR in Variable
count_crr_1=crr_1.Asset_id.sort_values().count()
count_crr_2=crr_2.Asset_id.sort_values().count()

#Compare both CRR count to check whether both are matched or not, If not matched then display message and generate one missing report
def compare_count(c1,c2):
    if c1 != c2:
        print('crr_1 is not matched with crr_2')
        unmatched_contract=crr_1.merge(crr_2,on='contracts',how='outer')
        unmatched_contract=unmatched_contract.rename(columns={'Asset_id_x':'CRR_1_Asset_id','underlying_x':'CRR_1_underlying','Asset_id_y':'CRR_2_Asset_id','underlying_y':'CRR_2_underlying'})
        unmatched_contract.to_excel(r'D:\Operations\New_CRR_comparision\missing_data.xlsx',sheet_name='missing_data',na_rep='NAN')

compare_count(count_crr_1,count_crr_2)

# Merge CRR_1 with CRR_2 using left join
Marged_CRR =crr_1.merge(crr_2,on='contracts',how='outer')

Marged_CRR['difference'] = (Marged_CRR.underlying_y - Marged_CRR.underlying_x)
Marged_CRR['percentage %'] = ((Marged_CRR['difference']/Marged_CRR.underlying_x)*100).round(2)

Marged_CRR=Marged_CRR.rename(columns={'Asset_id_x':'CRR_1_Asset_id','underlying_x':'CRR_1_underlying','Asset_id_y':'CRR_2_Asset_id','underlying_y':'CRR_2_underlying'})

# Replace all 0 values with NAN
final=Marged_CRR.replace(0,np.nan)

# Drop all NAN values
final_output=final.dropna(axis=0).reset_index(drop=1)

def color_postive_value(val):
    color = 'red' if val >= 5 or  val <=-5 else 'black'
    return 'color :%s' % color

final_output=final_output.style.applymap(color_postive_value,subset=['percentage %'])

final_output.to_excel(r'D:\Operations\New_CRR_comparision\CRR_comparison.xlsx',sheet_name='CRR_COMP',na_rep='')