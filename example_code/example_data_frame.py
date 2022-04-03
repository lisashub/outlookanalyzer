import pandas as pd
import numpy as np
import dataframe_image as dfi
    
# Initialise data to lists. 
data = [{'Geeks': 'dataframe', 'For': 'using', 'geeks': 'list'},
        {'Geeks':10, 'For': 20, 'geeks': 30}] 
    
# Creates DataFrame. 
df = pd.DataFrame(data) 
    
# Print the data 
df 

df_styled = df.style.background_gradient() #adding a gradient based on values in cell
dfi.export(df_styled,"mytable.png")

# df = pd.DataFrame(np.random.randn(6, 6), columns=list('ABCDEF'))