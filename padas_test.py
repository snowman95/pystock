import pandas as pd
import numpy as np
df = pd.DataFrame(np.array([[1, 2, 3], # → 이게 하나의 행임. Excel의 가로축 
                            [4, 5, 6], 
                            [7, 8, 9]]))
                            #columns=['A', 'B', 'C'],
                            #index=['ONE', 'TWO', 'THREE'])

print('''df''')
print(df)
print('\n')

#df = pd.DataFrame({'A': [1, 2, 3],
#                   'B': [4, 5, 6],
#                   'C': [7, 8, 9]},
#                   index=['ONE', 'TWO', 'THREE'])

#print('df')
#print(df)
#print('\n')
#
#print('''df['A']''')
#print(df['A'])
#print('\n')
#
#print('df.values')
#print(df.values)
#print('\n')
#
#print('''df['A'].values''')
#print(df['A'].values)
#print('\n')
#
#print('''df['A']['ONE']''')
#print(df['A']['ONE'])
#print('\n')
#
#print('''df.loc['ONE']''')
#print(df.loc['ONE'])

