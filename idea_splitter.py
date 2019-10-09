import pandas as pd
import os
import re

def get_domain(description):
    '''Returns domain of the description'''
    description = description.lower()
    print(description)
    domains = re.findall("^domain[.: -]*[a-z]*",description)
    print(domains)
    if (not domains):
        return 0
    return domains[0]

def write_record_to_file(domain, record):
    '''Writes the given record to the appropriate file'''
    #TODO

data = pd.read_csv("hashcode-idea-submissions.csv")
data = pd.DataFrame(data)

for ind,row in data.iterrows(): #iterate over all records
    domain = get_domain(row['Description'])
    if (not pd.isna(row['Attachment'])):
        print("wgetting attach", row["Attachment"])
        #os.system("wget -O "+"./domains/"+domain+"/"+row['Idea title']+".pdf "+row['Attachment'])
