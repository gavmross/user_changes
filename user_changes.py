import pandas as pd
xls = pd.ExcelFile('/Users/gavmross/Desktop/cnct/profile_changes/profile_data/profile_changes_10_20.xlsx')
df_pastWeek = pd.read_excel(xls, "cw", index_col=0).fillna(0)
df_currentWeek = pd.read_excel("/Users/gavmross/Desktop/cnct/profile_changes/cw.xlsx").fillna(0)

cols = df_pastWeek.columns
# ++++++++++++ Deleted Users = users in last week data that are not in this week data +++++++++++++++
df_del = df_pastWeek.loc[~df_pastWeek['userId'].isin(df_currentWeek['userId'])]

# +++++++++++ New Users = users in this weeks data that are not in last weeks data ++++++++++++++

df_new = df_currentWeek.loc[~df_currentWeek['userId'].isin(df_pastWeek['userId'])]

#++++++++++ Find Users with unfiniished profiles (in cw) +++++++++++++
df_unfinished = df_currentWeek.loc[df_currentWeek['firstName'] == 0]
df_unfinished.reset_index(inplace=True, drop=True)

# abs(new users - deleted users) = abs(total users this week - total users last week)

def check_counts(new, deleted, cw, pw):
    return abs(new-deleted) == abs(cw-pw)

print("Does math check out?", check_counts(df_new.shape[0], df_del.shape[0], df_currentWeek.shape[0], df_pastWeek.shape[0]))

# To find differences in data, need to compare only users in both weeks
# Need to take out new users from this week (since they are not in last week)
# Need to take out deleted users form last week (since they are not in this week)
df_cw_compare = df_currentWeek.loc[~df_currentWeek['userId'].isin(df_new['userId'])]
df_pw_compare = df_pastWeek.loc[~df_pastWeek['userId'].isin(df_del['userId'])]
df_cw_compare.reset_index(inplace=True, drop=True)
df_pw_compare.reset_index(inplace=True, drop=True)
print(df_cw_compare.shape)
print(df_pw_compare.shape)

#strip all strings so they can compare properly 
def trim_all_columns(df):
    """
    Trim whitespace from ends of each value across all series in dataframe
    """
    trim_strings = lambda x: x.strip() if isinstance(x, str) else x
    return df.applymap(trim_strings)

df_cw_compare = trim_all_columns(df_cw_compare)
df_pw_compare = trim_all_columns(df_pw_compare)

df_differences = df_pw_compare.compare(df_cw_compare)


#=================================================================
fname = 'profile_changes_10_27.xlsx'
with pd.ExcelWriter(fname, engine='xlsxwriter') as writer:
    df_differences.to_excel(writer, sheet_name='differences')
    df_cw_compare.to_excel(writer, sheet_name='cw compare')
    df_pw_compare.to_excel(writer, sheet_name='pw compare')
    df_unfinished.to_excel(writer, sheet_name='unfinished')
    df_new.to_excel(writer, sheet_name='new users')
    df_del.to_excel(writer, sheet_name='del users')
    df_currentWeek.to_excel(writer, sheet_name='cw')
    df_pastWeek.to_excel(writer, sheet_name='pw')
