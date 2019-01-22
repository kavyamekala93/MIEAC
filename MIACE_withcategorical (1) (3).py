import sys
import pandas as pd
import numpy as np

def miaec_column(c):
    #Check column is numerical or categorical data
    number_flag = np.issubdtype(c.dtype, np.number)

    # Mark Missing Location
    missing_index = c[pd.isna(c)].index
    final_mean = c.mean()
    # Perform mean imputation as a chain of evidence for each missing
    candidates = []
    for i in missing_index:
        if i == 0:
            for j in c.index:
                if pd.notna(c[j]):
                    c[i] = c[j]
                    break
        else:
            if pd.isna(c[i]):
                c_chain = c[0:i]
                if number_flag:
                    c[i] = c_chain.mean()
                else:
                    c[i] = c_chain.value_counts().idxmax()
        candidates.append(c[i])
    #Find the candidate with min std
    if len(candidates) > 0:
        if number_flag:
            choose_candidate = candidates[-1]
            
            print(final_mean)
            for candidate in candidates:
                if(abs(candidate - final_mean) < abs(choose_candidate - final_mean)):
                    choose_candidate = candidate
        else:
            for candidate_max in c.value_counts().index:
                if candidate_max in candidates:
                    choose_candidate = candidate_max
                    break
        #print 'choose ', choose_candidate

    #All missing location is filled with chosen candidate
        c[missing_index] = choose_candidate
    return c

def main(argv):
    print ('Working with file: ', r'K:\Data_Mining\Project\SourceCode\Comp20.xlsx')
    df = pd.read_excel(r'K:\Data_Mining\Project\SourceCode\Comp20.xlsx', header=None)
    df_out = df.apply(miaec_column, axis = 0)
    df_out.to_excel(r'K:\Data_Mining\Project\SourceCode\Comp20' + '-out.xlsx', header=None, index=None)
    print ('Finish output file ' + r'K:\Data_Mining\Project\SourceCode\Comp20' + '-out.xlsx')
    
if __name__ == "__main__":
   main(sys.argv[1:])


