import pandas as pd

data = [['tom', 10], ['nick', 15], ['juli', 14]]
  
# Create the pandas DataFrame
df = pd.DataFrame(data, columns = ['Name', 'Age'])
  
df.to_csv("C:\Users\ialonso\test.csv")
