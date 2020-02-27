### The code snippets are examples of how to extract data from spreadsheets


```python
import pandas as pd
import numpy as np
import janitor
```

### Case 4: Pivot Table - Single Headers


```python
df = (pd.read_excel('worked-examples.xlsx',
                     sheet_name='pivot-hierarchy',
                     header=None)
        )

df
```




<div>
<style scoped>
    .dataframe tbody tr th:only-of-type {
        vertical-align: middle;
    }

    .dataframe tbody tr th {
        vertical-align: top;
    }

    .dataframe thead th {
        text-align: right;
    }
</style>
<table border="1" class="dataframe">
  <thead>
    <tr style="text-align: right;">
      <th></th>
      <th>0</th>
      <th>1</th>
      <th>2</th>
      <th>3</th>
    </tr>
  </thead>
  <tbody>
    <tr>
      <th>0</th>
      <td>NaN</td>
      <td>NaN</td>
      <td>NaN</td>
      <td>NaN</td>
    </tr>
    <tr>
      <th>1</th>
      <td>NaN</td>
      <td>NaN</td>
      <td>Matilda</td>
      <td>Nicholas</td>
    </tr>
    <tr>
      <th>2</th>
      <td>NaN</td>
      <td>Humanities</td>
      <td>NaN</td>
      <td>NaN</td>
    </tr>
    <tr>
      <th>3</th>
      <td>NaN</td>
      <td>Classics</td>
      <td>1</td>
      <td>3</td>
    </tr>
    <tr>
      <th>4</th>
      <td>NaN</td>
      <td>History</td>
      <td>3</td>
      <td>5</td>
    </tr>
    <tr>
      <th>5</th>
      <td>NaN</td>
      <td>Performance</td>
      <td>NaN</td>
      <td>NaN</td>
    </tr>
    <tr>
      <th>6</th>
      <td>NaN</td>
      <td>Music</td>
      <td>5</td>
      <td>9</td>
    </tr>
    <tr>
      <th>7</th>
      <td>NaN</td>
      <td>Drama</td>
      <td>7</td>
      <td>12</td>
    </tr>
  </tbody>
</table>
</div>




```python
# Helper Functions

def extract_field_col(df,col,ref,new_col):
    '''
    Creates the field column and returns a dataframe.
    '''
    cond = df[col].isna()
    df[new_col] = np.where(cond,df[ref],np.nan)
    return df

def fill_col(df,col):
    '''
    Fills null values forward in a column
    and returns a dataframe.
    '''
    df[col] = df[col].ffill()
    return df

def remove_rows(df):
    '''
    Compares the first and last column,
    checks if any rows have the same text,
    removes the rows,
    and returns a dataframe.
    '''
    cond = df.iloc[:,0].eq(df.iloc[:,-1])
    df = df.loc[~cond]
    return df
```


```python
#Tidy data processing

(df
 #remove_empty is a function in pyjanitor
 #it removes completely empty rows and columns
 .remove_empty()
 .pipe(extract_field_col,2,1,'Field')
 .pipe(fill_col,'Field')
 .pipe(remove_rows)
 .fillna({1:'Subject',
         'Field':'Field'})
 .row_to_names(row_number=0, remove_row=True)
 .melt(id_vars=['Subject','Field'],
       value_name='Score',
       var_name='Student')
)
```




<div>
<style scoped>
    .dataframe tbody tr th:only-of-type {
        vertical-align: middle;
    }

    .dataframe tbody tr th {
        vertical-align: top;
    }

    .dataframe thead th {
        text-align: right;
    }
</style>
<table border="1" class="dataframe">
  <thead>
    <tr style="text-align: right;">
      <th></th>
      <th>Subject</th>
      <th>Field</th>
      <th>Student</th>
      <th>Score</th>
    </tr>
  </thead>
  <tbody>
    <tr>
      <th>0</th>
      <td>Classics</td>
      <td>Humanities</td>
      <td>Matilda</td>
      <td>1</td>
    </tr>
    <tr>
      <th>1</th>
      <td>History</td>
      <td>Humanities</td>
      <td>Matilda</td>
      <td>3</td>
    </tr>
    <tr>
      <th>2</th>
      <td>Music</td>
      <td>Performance</td>
      <td>Matilda</td>
      <td>5</td>
    </tr>
    <tr>
      <th>3</th>
      <td>Drama</td>
      <td>Performance</td>
      <td>Matilda</td>
      <td>7</td>
    </tr>
    <tr>
      <th>4</th>
      <td>Classics</td>
      <td>Humanities</td>
      <td>Nicholas</td>
      <td>3</td>
    </tr>
    <tr>
      <th>5</th>
      <td>History</td>
      <td>Humanities</td>
      <td>Nicholas</td>
      <td>5</td>
    </tr>
    <tr>
      <th>6</th>
      <td>Music</td>
      <td>Performance</td>
      <td>Nicholas</td>
      <td>9</td>
    </tr>
    <tr>
      <th>7</th>
      <td>Drama</td>
      <td>Performance</td>
      <td>Nicholas</td>
      <td>12</td>
    </tr>
  </tbody>
</table>
</div>



### Case 5 : Implied Multiples


```python
#Helper Function

def extract_col_grade(df,col,new_col,ref,text):
    
    cond = df[col]==text
    
    df[new_col] = np.where(cond,df[ref],np.nan)
    
    return df
```


```python
#actual code

df = (pd.read_excel('worked-examples.xlsx',
                    sheet_name='implied-multiples',
                    header=None)
      .ffill(axis=1)
      .T
      .fillna('Field')
      #row_to_names is a function from the pyjanitor package
      .row_to_names(row_number = 0,
                    remove_row = True)
      .melt(id_vars=['Field','Name'],
            var_name='Student',
            value_name = 'Score')
      .pipe(extract_col_grade,'Name','Grade','Score','Grade')
      .bfill()
      .query('Name != "Grade"')
      .reset_index(drop = True)
      .rename(columns={"Name":'Subject'})
     )
```

### Case 6: Table - Centre-aligned headers


```python
#openpyxl gets us the coordinates for the borders

from more_itertools import windowed
from openpyxl import load_workbook

wb = load_workbook(filename='worked-examples.xlsx')

ws = wb['pivot-centre-aligned']

cols = set()
rows = set()

for line in ws:
    for cell in line:
        if cell.border.right.style:
            cols.add(cell.column)
        elif cell.border.bottom.style:
            rows.add(cell.row)
            
cols = sorted(cols)
rows = sorted(rows)

col_count = len(cols) - 1 # 'box' count
row_count = len(rows) - 1

rows = list(windowed(rows, row_count))
cols = list(windowed(cols, col_count))

print(rows,cols)
```

    [(3, 8), (8, 11)] [(3, 6), (6, 10)]



```python
df = (pd.read_excel('worked-examples.xlsx',
                    sheet_name='pivot-centre-aligned',
                    header=None)
     )

df
```




<div>
<style scoped>
    .dataframe tbody tr th:only-of-type {
        vertical-align: middle;
    }

    .dataframe tbody tr th {
        vertical-align: top;
    }

    .dataframe thead th {
        text-align: right;
    }
</style>
<table border="1" class="dataframe">
  <thead>
    <tr style="text-align: right;">
      <th></th>
      <th>0</th>
      <th>1</th>
      <th>2</th>
      <th>3</th>
      <th>4</th>
      <th>5</th>
      <th>6</th>
      <th>7</th>
      <th>8</th>
      <th>9</th>
    </tr>
  </thead>
  <tbody>
    <tr>
      <th>0</th>
      <td>NaN</td>
      <td>NaN</td>
      <td>NaN</td>
      <td>NaN</td>
      <td>NaN</td>
      <td>NaN</td>
      <td>NaN</td>
      <td>NaN</td>
      <td>NaN</td>
      <td>NaN</td>
    </tr>
    <tr>
      <th>1</th>
      <td>NaN</td>
      <td>NaN</td>
      <td>NaN</td>
      <td>NaN</td>
      <td>Female</td>
      <td>NaN</td>
      <td>NaN</td>
      <td>NaN</td>
      <td>Male</td>
      <td>NaN</td>
    </tr>
    <tr>
      <th>2</th>
      <td>NaN</td>
      <td>NaN</td>
      <td>NaN</td>
      <td>Leah</td>
      <td>Matilda</td>
      <td>Olivia</td>
      <td>Lenny</td>
      <td>Max</td>
      <td>Nicholas</td>
      <td>Paul</td>
    </tr>
    <tr>
      <th>3</th>
      <td>NaN</td>
      <td>NaN</td>
      <td>Classics</td>
      <td>3</td>
      <td>1</td>
      <td>2</td>
      <td>4</td>
      <td>3</td>
      <td>3</td>
      <td>0</td>
    </tr>
    <tr>
      <th>4</th>
      <td>NaN</td>
      <td>Humanities</td>
      <td>History</td>
      <td>8</td>
      <td>3</td>
      <td>4</td>
      <td>7</td>
      <td>5</td>
      <td>5</td>
      <td>1</td>
    </tr>
    <tr>
      <th>5</th>
      <td>NaN</td>
      <td>NaN</td>
      <td>Literature</td>
      <td>1</td>
      <td>1</td>
      <td>9</td>
      <td>3</td>
      <td>12</td>
      <td>7</td>
      <td>5</td>
    </tr>
    <tr>
      <th>6</th>
      <td>NaN</td>
      <td>NaN</td>
      <td>Philosophy</td>
      <td>5</td>
      <td>10</td>
      <td>10</td>
      <td>8</td>
      <td>2</td>
      <td>5</td>
      <td>12</td>
    </tr>
    <tr>
      <th>7</th>
      <td>NaN</td>
      <td>NaN</td>
      <td>Languages</td>
      <td>5</td>
      <td>4</td>
      <td>5</td>
      <td>9</td>
      <td>8</td>
      <td>3</td>
      <td>8</td>
    </tr>
    <tr>
      <th>8</th>
      <td>NaN</td>
      <td>NaN</td>
      <td>Music</td>
      <td>4</td>
      <td>10</td>
      <td>10</td>
      <td>2</td>
      <td>4</td>
      <td>5</td>
      <td>6</td>
    </tr>
    <tr>
      <th>9</th>
      <td>NaN</td>
      <td>Performance</td>
      <td>Dance</td>
      <td>4</td>
      <td>5</td>
      <td>6</td>
      <td>4</td>
      <td>12</td>
      <td>9</td>
      <td>2</td>
    </tr>
    <tr>
      <th>10</th>
      <td>NaN</td>
      <td>NaN</td>
      <td>Drama</td>
      <td>2</td>
      <td>7</td>
      <td>8</td>
      <td>6</td>
      <td>1</td>
      <td>12</td>
      <td>3</td>
    </tr>
  </tbody>
</table>
</div>




```python
#extract gender and names

gender_and_names = [df.copy()
                    .iloc[:,start:end]
                    .dropna(how='all')
                    .ffill(axis=1)
                    .bfill(axis=1)
                    for start,end
                    in cols
                   ]

#merge extracts

gender = pd.concat(gender_and_names,axis=1)

#extract fields

Fields = [df.copy()
          .iloc[start:end]
          .dropna(how='all',axis=1)
          .ffill()
          .bfill()
          .iloc[:,:2] #only interested in first two columns
          for start,end in rows
         ]

#merge extracts

fields = pd.concat(Fields)
```


```python
#Helper Functions

def rename_col(df):
    '''
    Merges first two rows of dataframe,
    assigns aggregation as the new column names,
    drops the first two rows from the dataframe
    and returns a dataframe
    '''
    
    df.columns = df.iloc[:2].add(',').sum().str.strip(',')
    
    df = df.iloc[2:]
     
    return df

def create_col(df,col,new_col):
    '''
    Extracts new_col from col;
    splits col into two,
    assigns first part of the split to new_col,
    second part to col
    and returns a dataframe
    '''
    
    df[new_col] = df[col].str.split(',').str[0]
    
    df[col] = df[col].str.split(',').str[-1]
    
    return df
```


```python
outcome = (pd.concat([fields,gender],axis=1)
           .fillna(value={1:'Field',
                          2:'Subject'},
                   limit=1)
           .fillna('')
           .pipe(rename_col)
           .melt(id_vars=['Field','Subject'],
                 var_name='Student',
                 value_name='Score')
           .pipe(create_col,'Student','Gender')
            )

outcome
```




<div>
<style scoped>
    .dataframe tbody tr th:only-of-type {
        vertical-align: middle;
    }

    .dataframe tbody tr th {
        vertical-align: top;
    }

    .dataframe thead th {
        text-align: right;
    }
</style>
<table border="1" class="dataframe">
  <thead>
    <tr style="text-align: right;">
      <th></th>
      <th>Field</th>
      <th>Subject</th>
      <th>Student</th>
      <th>Score</th>
      <th>Gender</th>
    </tr>
  </thead>
  <tbody>
    <tr>
      <th>0</th>
      <td>Humanities</td>
      <td>Classics</td>
      <td>Leah</td>
      <td>3</td>
      <td>Female</td>
    </tr>
    <tr>
      <th>1</th>
      <td>Humanities</td>
      <td>History</td>
      <td>Leah</td>
      <td>8</td>
      <td>Female</td>
    </tr>
    <tr>
      <th>2</th>
      <td>Humanities</td>
      <td>Literature</td>
      <td>Leah</td>
      <td>1</td>
      <td>Female</td>
    </tr>
    <tr>
      <th>3</th>
      <td>Humanities</td>
      <td>Philosophy</td>
      <td>Leah</td>
      <td>5</td>
      <td>Female</td>
    </tr>
    <tr>
      <th>4</th>
      <td>Humanities</td>
      <td>Languages</td>
      <td>Leah</td>
      <td>5</td>
      <td>Female</td>
    </tr>
    <tr>
      <th>5</th>
      <td>Performance</td>
      <td>Music</td>
      <td>Leah</td>
      <td>4</td>
      <td>Female</td>
    </tr>
    <tr>
      <th>6</th>
      <td>Performance</td>
      <td>Dance</td>
      <td>Leah</td>
      <td>4</td>
      <td>Female</td>
    </tr>
    <tr>
      <th>7</th>
      <td>Performance</td>
      <td>Drama</td>
      <td>Leah</td>
      <td>2</td>
      <td>Female</td>
    </tr>
    <tr>
      <th>8</th>
      <td>Humanities</td>
      <td>Classics</td>
      <td>Matilda</td>
      <td>1</td>
      <td>Female</td>
    </tr>
    <tr>
      <th>9</th>
      <td>Humanities</td>
      <td>History</td>
      <td>Matilda</td>
      <td>3</td>
      <td>Female</td>
    </tr>
    <tr>
      <th>10</th>
      <td>Humanities</td>
      <td>Literature</td>
      <td>Matilda</td>
      <td>1</td>
      <td>Female</td>
    </tr>
    <tr>
      <th>11</th>
      <td>Humanities</td>
      <td>Philosophy</td>
      <td>Matilda</td>
      <td>10</td>
      <td>Female</td>
    </tr>
    <tr>
      <th>12</th>
      <td>Humanities</td>
      <td>Languages</td>
      <td>Matilda</td>
      <td>4</td>
      <td>Female</td>
    </tr>
    <tr>
      <th>13</th>
      <td>Performance</td>
      <td>Music</td>
      <td>Matilda</td>
      <td>10</td>
      <td>Female</td>
    </tr>
    <tr>
      <th>14</th>
      <td>Performance</td>
      <td>Dance</td>
      <td>Matilda</td>
      <td>5</td>
      <td>Female</td>
    </tr>
    <tr>
      <th>15</th>
      <td>Performance</td>
      <td>Drama</td>
      <td>Matilda</td>
      <td>7</td>
      <td>Female</td>
    </tr>
    <tr>
      <th>16</th>
      <td>Humanities</td>
      <td>Classics</td>
      <td>Olivia</td>
      <td>2</td>
      <td>Female</td>
    </tr>
    <tr>
      <th>17</th>
      <td>Humanities</td>
      <td>History</td>
      <td>Olivia</td>
      <td>4</td>
      <td>Female</td>
    </tr>
    <tr>
      <th>18</th>
      <td>Humanities</td>
      <td>Literature</td>
      <td>Olivia</td>
      <td>9</td>
      <td>Female</td>
    </tr>
    <tr>
      <th>19</th>
      <td>Humanities</td>
      <td>Philosophy</td>
      <td>Olivia</td>
      <td>10</td>
      <td>Female</td>
    </tr>
    <tr>
      <th>20</th>
      <td>Humanities</td>
      <td>Languages</td>
      <td>Olivia</td>
      <td>5</td>
      <td>Female</td>
    </tr>
    <tr>
      <th>21</th>
      <td>Performance</td>
      <td>Music</td>
      <td>Olivia</td>
      <td>10</td>
      <td>Female</td>
    </tr>
    <tr>
      <th>22</th>
      <td>Performance</td>
      <td>Dance</td>
      <td>Olivia</td>
      <td>6</td>
      <td>Female</td>
    </tr>
    <tr>
      <th>23</th>
      <td>Performance</td>
      <td>Drama</td>
      <td>Olivia</td>
      <td>8</td>
      <td>Female</td>
    </tr>
    <tr>
      <th>24</th>
      <td>Humanities</td>
      <td>Classics</td>
      <td>Lenny</td>
      <td>4</td>
      <td>Male</td>
    </tr>
    <tr>
      <th>25</th>
      <td>Humanities</td>
      <td>History</td>
      <td>Lenny</td>
      <td>7</td>
      <td>Male</td>
    </tr>
    <tr>
      <th>26</th>
      <td>Humanities</td>
      <td>Literature</td>
      <td>Lenny</td>
      <td>3</td>
      <td>Male</td>
    </tr>
    <tr>
      <th>27</th>
      <td>Humanities</td>
      <td>Philosophy</td>
      <td>Lenny</td>
      <td>8</td>
      <td>Male</td>
    </tr>
    <tr>
      <th>28</th>
      <td>Humanities</td>
      <td>Languages</td>
      <td>Lenny</td>
      <td>9</td>
      <td>Male</td>
    </tr>
    <tr>
      <th>29</th>
      <td>Performance</td>
      <td>Music</td>
      <td>Lenny</td>
      <td>2</td>
      <td>Male</td>
    </tr>
    <tr>
      <th>30</th>
      <td>Performance</td>
      <td>Dance</td>
      <td>Lenny</td>
      <td>4</td>
      <td>Male</td>
    </tr>
    <tr>
      <th>31</th>
      <td>Performance</td>
      <td>Drama</td>
      <td>Lenny</td>
      <td>6</td>
      <td>Male</td>
    </tr>
    <tr>
      <th>32</th>
      <td>Humanities</td>
      <td>Classics</td>
      <td>Max</td>
      <td>3</td>
      <td>Male</td>
    </tr>
    <tr>
      <th>33</th>
      <td>Humanities</td>
      <td>History</td>
      <td>Max</td>
      <td>5</td>
      <td>Male</td>
    </tr>
    <tr>
      <th>34</th>
      <td>Humanities</td>
      <td>Literature</td>
      <td>Max</td>
      <td>12</td>
      <td>Male</td>
    </tr>
    <tr>
      <th>35</th>
      <td>Humanities</td>
      <td>Philosophy</td>
      <td>Max</td>
      <td>2</td>
      <td>Male</td>
    </tr>
    <tr>
      <th>36</th>
      <td>Humanities</td>
      <td>Languages</td>
      <td>Max</td>
      <td>8</td>
      <td>Male</td>
    </tr>
    <tr>
      <th>37</th>
      <td>Performance</td>
      <td>Music</td>
      <td>Max</td>
      <td>4</td>
      <td>Male</td>
    </tr>
    <tr>
      <th>38</th>
      <td>Performance</td>
      <td>Dance</td>
      <td>Max</td>
      <td>12</td>
      <td>Male</td>
    </tr>
    <tr>
      <th>39</th>
      <td>Performance</td>
      <td>Drama</td>
      <td>Max</td>
      <td>1</td>
      <td>Male</td>
    </tr>
    <tr>
      <th>40</th>
      <td>Humanities</td>
      <td>Classics</td>
      <td>Nicholas</td>
      <td>3</td>
      <td>Male</td>
    </tr>
    <tr>
      <th>41</th>
      <td>Humanities</td>
      <td>History</td>
      <td>Nicholas</td>
      <td>5</td>
      <td>Male</td>
    </tr>
    <tr>
      <th>42</th>
      <td>Humanities</td>
      <td>Literature</td>
      <td>Nicholas</td>
      <td>7</td>
      <td>Male</td>
    </tr>
    <tr>
      <th>43</th>
      <td>Humanities</td>
      <td>Philosophy</td>
      <td>Nicholas</td>
      <td>5</td>
      <td>Male</td>
    </tr>
    <tr>
      <th>44</th>
      <td>Humanities</td>
      <td>Languages</td>
      <td>Nicholas</td>
      <td>3</td>
      <td>Male</td>
    </tr>
    <tr>
      <th>45</th>
      <td>Performance</td>
      <td>Music</td>
      <td>Nicholas</td>
      <td>5</td>
      <td>Male</td>
    </tr>
    <tr>
      <th>46</th>
      <td>Performance</td>
      <td>Dance</td>
      <td>Nicholas</td>
      <td>9</td>
      <td>Male</td>
    </tr>
    <tr>
      <th>47</th>
      <td>Performance</td>
      <td>Drama</td>
      <td>Nicholas</td>
      <td>12</td>
      <td>Male</td>
    </tr>
    <tr>
      <th>48</th>
      <td>Humanities</td>
      <td>Classics</td>
      <td>Paul</td>
      <td>0</td>
      <td>Male</td>
    </tr>
    <tr>
      <th>49</th>
      <td>Humanities</td>
      <td>History</td>
      <td>Paul</td>
      <td>1</td>
      <td>Male</td>
    </tr>
    <tr>
      <th>50</th>
      <td>Humanities</td>
      <td>Literature</td>
      <td>Paul</td>
      <td>5</td>
      <td>Male</td>
    </tr>
    <tr>
      <th>51</th>
      <td>Humanities</td>
      <td>Philosophy</td>
      <td>Paul</td>
      <td>12</td>
      <td>Male</td>
    </tr>
    <tr>
      <th>52</th>
      <td>Humanities</td>
      <td>Languages</td>
      <td>Paul</td>
      <td>8</td>
      <td>Male</td>
    </tr>
    <tr>
      <th>53</th>
      <td>Performance</td>
      <td>Music</td>
      <td>Paul</td>
      <td>6</td>
      <td>Male</td>
    </tr>
    <tr>
      <th>54</th>
      <td>Performance</td>
      <td>Dance</td>
      <td>Paul</td>
      <td>2</td>
      <td>Male</td>
    </tr>
    <tr>
      <th>55</th>
      <td>Performance</td>
      <td>Drama</td>
      <td>Paul</td>
      <td>3</td>
      <td>Male</td>
    </tr>
  </tbody>
</table>
</div>


