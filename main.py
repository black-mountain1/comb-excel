import pandas as pd
import glob
import argparse


def parse_args():
    parser = argparse.ArgumentParser(description='Tool to combine monthly sales reports.')
    parser.add_argument('data_dir', help='Directory where the sales data is stored')
    parser.add_argument('output_dir', help='Directory where the summary report will be stored')
    parser.add_argument('cust_file', help='Customer account status file')
    parser.add_argument('-d', help='Start date to include')

    args = parser.parse_args()
    return args


def read_sales(dir):
    """
    Reads sales data files in the format sales-*-*.xlsx from the directory.
    :param dir: Location of sales data files
    :return: DataFrame of combined sales data
    """
    all_data = pd.DataFrame()
    for file in glob.glob(dir + '/sales-*-*.xlsx'):
        df = pd.read_excel(file)
        all_data = all_data.append(df, ignore_index=True)

    return all_data


def read_cust_file(dir):
    """
    Reads the customer status file 'customer-status.xlsx in the directory.
    :param dir: Location of the customer file
    :return: DataFrame of the customer data
    """
    df = pd.read_excel(dir + '/customer-status.xlsx')
    return df


def process_data(all_data, cust_data):
    """
    Processes the sales data by joining it with the customer data and sorting on status category.
    :param all_data: DataFrame of the sales data
    :param cust_data: DataFrame of the customer data
    :return: DataFrame of the combined sales and customer data
    """
    for col in all_data.columns:
        pct_missing = all_data[col].isna().sum()
        print(f"{col} column has {pct_missing} missing values")

    all_data['date'] = pd.to_datetime(all_data['date'])

    all_data_st = all_data.merge(cust_data, how='left')

    all_data_st.status.fillna('bronze', inplace=True)
    all_data_st['status'] = all_data_st['status'].astype('category')
    all_data_st['status'].cat.set_categories(['gold', 'silver', 'bronze'], inplace=True)
    all_data_st = all_data_st.sort_values(by='status')

    return all_data_st


def write_summary(dir, df):
    """
    Writes the DataFrame to the directory as summary_report.xlsx
    :param dir: Directory to save the report
    :param df: DataFrame of the summary report to be stored
    :return: No return
    """
    output_file = '/summary_report.xlsx'

    with pd.ExcelWriter(dir + output_file) as writer:
        df.to_excel(writer)


if __name__ == '__main__':
    args = parse_args()
    print('Reading sales data')
    sales_data = read_sales(args.data_dir)
    print('Reading customer data')
    cust = read_cust_file(args.cust_file)
    print('Processing data')
    processed_data = process_data(sales_data, cust)
    print('Saving summary report')
    write_summary(args.output_dir, processed_data)
    print('Script complete')
