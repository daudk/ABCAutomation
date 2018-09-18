import pandas as pd
import datetime as dt
import os

def sox_test_samples(ctrl_numb,samp_size,pickle_path):
    start = dt.datetime.now()
    data = pd.read_pickle(pickle_path)
    data_rel = data[(data['Dashed Invoice'] == 'Not Dashed') & (data['Expense Report Status'] == 'Not Expense Report')
                    & (data['Cost Purchase Status'] == 'Not Cost Purchase') & (data['PO Number'] == 'No PO number')]
    writer = pd.ExcelWriter(os.path.dirname(pickle_path) +'/{0}.xlsx'.format(ctrl_numb), engine='xlsxwriter')


    if (ctrl_numb=="15.2.3.02 F&A"):
        pop = data_rel[data_rel['division'] == 'Admin'].reset_index().drop('index',axis=1)
    elif (ctrl_numb=="15.2.1.08 BISG"):
        pop=data_rel[data_rel['division'] == 'BISG'].reset_index().drop('index',axis=1)
    elif (ctrl_numb=="15.2.1.10 ITCG"):
        pop = data_rel[data_rel['division'] == 'ITCG'].reset_index().drop('index',axis=1)
    elif (ctrl_numb=="15.2.1.07 LAG"):
        pop = data_rel[data_rel['division'] == 'LAG'].reset_index().drop('index',axis=1)
    else:
        pop = data_rel


    samp = pop.sample(int(samp_size))
    samp.reset_index(inplace=True)
    samp['index'] = samp['index'] + 2
    samp.rename(columns={'index': 'Old Row Number'}, inplace=True)

    pop.to_excel(writer, sheet_name='{0} population'.format(ctrl_numb), index=False)
    samp.to_excel(writer, sheet_name='{0} sample'.format(ctrl_numb), index=False)

    writer.save()

    print("Total seconds: ", (dt.datetime.now() - start).total_seconds())