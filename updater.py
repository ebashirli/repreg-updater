from datetime import datetime
import pandas as pd
import pyodbc

path = r"C:\Users\Elvin.Bashirli\Desktop\MC\New folder"



loc_systems = path + r"\systems (cms).xlsx"
loc_subsystems = path + r"\subsystems.xlsx"
loc_packages = path + r"\packages.xlsx"

df_projects   = pd.DataFrame(['Drilling', 'Living Quarters', 'Topsides'], columns=['Project'])
df_systems    = pd.DataFrame(pd.read_excel(loc_systems), columns=['Project', 'System'])
df_subsystems = pd.DataFrame(pd.read_excel(loc_subsystems), columns=['Project', 'System', 'Subsystem'])
df_packages   = pd.DataFrame(pd.read_excel(loc_packages), columns=['Project', 'System', 'Subsystem', 'Package'])

for i in range(len(df_packages)):
    pack_subs = df_packages.loc[i, 'Subsystem']
    pack_proj = df_packages.loc[i, 'Project']
    tf = (df_subsystems['Subsystem'] == pack_subs) & (df_subsystems['Project'] == pack_proj)
    df_packages.loc[i, 'System'] = list(df_subsystems.loc[tf, 'System'])[0]

for i in range(df_projects.shape[0]):
    df_systems.loc[df_systems['Project'] == df_projects.loc[i, 'Project'], 'Project'] = i
    df_subsystems.loc[df_subsystems['Project'] == df_projects.loc[i, 'Project'], 'Project'] = i
    df_packages.loc[df_packages['Project'] == df_projects.loc[i, 'Project'], 'Project'] = i

for i in range(len(df_systems)):
    df_subsystems.loc[df_subsystems['System'] == df_systems['System'][i], 'System'] = i
    df_packages.loc[df_packages['System'] == df_systems['System'][i], 'System'] = i
    
for i in range(len(df_subsystems)):
    df_packages.loc[df_packages['Subsystem'] == df_subsystems['Subsystem'][i], 'Subsystem'] = i

tblDict = {
  'tblProjects': {
    'creatStr': 'id integer, project text',
    'columns': 'id, project',
    'df': df_projects.itertuples()
  },
  'tblSystems': {
    'creatStr': 'ID integer, projectID integer, system text',
    'columns': 'ID, projectID, system',
    'df': df_systems.itertuples()
  },
  'tblSubsystems': {
    'creatStr': 'id integer, projectID integer, systemID integer, subsystemNo text',
    'columns': 'id, projectID, systemID, subsystemNo',
    'df': df_subsystems.itertuples()
  },
  'tblPackages': {
    'creatStr': 'ID integer, projectID integer, systemID integer, subsystemID integer, package text',
    'columns': 'ID, projectID, systemID, subsystemID, package',
    'df': df_packages.itertuples()
  },
}

# db_path = r"C:\Users\Elvin.Bashirli\OneDrive - AZFEN J.V\Documents\RepReg\Registers\Registers.accdb"
db_path = "R:\Mechanical_Completion\Report & Register Interface\Registers\Registers.accdb"

conn = pyodbc.connect(r'Driver={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=' + db_path + ';PWD=EZtRr6N')
cursor = conn.cursor()

for key in tblDict.keys():
  cursor.execute(f'DROP TABLE {key};')
  cursor.execute(f"CREATE TABLE {key} ({tblDict[key]['creatStr']});")

  cursor.executemany(
      f"INSERT INTO {key} ({tblDict[key]['columns']}) VALUES ({', '.join('?'*(tblDict[key]['columns'].count(',') + 1))})",
      tblDict[key]['df'])

conn.commit()
print(datetime.now())
input(r"You are great, aren't you? ")