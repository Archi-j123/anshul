import pandas as pd

def add_final_result_column(df):

    def calculate_final_result(row):
        # Identify all columns ending with '_result'
        result_columns = [col for col in row.index if '_result' in col]
        # Check if all these columns have the value 'PASS'
        return 'PASS' if all(row[col] == 'PASS' for col in result_columns) else 'FAIL'
    
    df['Final_result'] = df.apply(calculate_final_result, axis=1)
    return df