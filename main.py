import pandas as pd
import glob
import getpass

user = getpass.getuser()
input_dir = f'/Users/{user}/Projects/data/comb-excel'

status = pd.read_excel(input_dir + '/customer-status.xlsx')
print(f'Shape of status: {status.shape}')

all_data = pd.DataFrame()
for file in glob.glob(input_dir + '/sales-*-*.xlsx'):
    df = pd.read_excel(file)
    all_data = all_data.append(df, ignore_index=True)


print("Shape of all data: ", all_data.shape)
print("Data types: ", all_data.dtypes)

for col in all_data.columns:
    pct_missing = all_data[col].isna().sum()
    print(f"{col} column has {pct_missing} missing values")

all_data['date'] = pd.to_datetime(all_data['date'])

all_data_st = all_data.merge(status, how='left')

all_data_st.status.fillna('bronze', inplace=True)
all_data_st['status'] = all_data_st['status'].astype('category')
all_data_st['status'].cat.set_categories(['gold', 'silver', 'bronze'], inplace=True)
all_data_st = all_data_st.sort_values(by='status')
