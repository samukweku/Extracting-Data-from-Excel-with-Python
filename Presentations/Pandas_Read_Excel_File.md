## The code snippets show how to read in Excel Files using Pandas

### Case 1: Read in an Excel File:


```python
import pandas as pd

filename = 'worked-examples.xlsx'

df = pd.read_excel(filename)

df.head()
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
      <th>Name</th>
      <th>Age</th>
    </tr>
  </thead>
  <tbody>
    <tr>
      <th>0</th>
      <td>Matilda</td>
      <td>1</td>
    </tr>
    <tr>
      <th>1</th>
      <td>Nicholas</td>
      <td>3</td>
    </tr>
    <tr>
      <th>2</th>
      <td>Olivia</td>
      <td>5</td>
    </tr>
  </tbody>
</table>
</div>



### Case 2 : Read in a particular sheet:


```python
df = pd.read_excel(filename, 
                   sheet_name = 'highlights')
df.head()
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
      <th>Age</th>
      <th>Height</th>
    </tr>
  </thead>
  <tbody>
    <tr>
      <th>0</th>
      <td>1</td>
      <td>2</td>
    </tr>
    <tr>
      <th>1</th>
      <td>3</td>
      <td>4</td>
    </tr>
    <tr>
      <th>2</th>
      <td>5</td>
      <td>6</td>
    </tr>
  </tbody>
</table>
</div>



### Case 3: Read in all the sheets:


```python
filename = 'EMT1626-Start.xlsx'

all_sheets = pd.read_excel(filename,sheet_name=None)

all_sheets
```




    {'Data(1)':         Date   Sales   Product
     0 2020-12-01   14.35      Quad
     1 2020-12-02  144.42  Sunshine
     2 2020-12-04  207.00      Quad
     3 2020-12-01  247.33  Sunshine,
     'Data(2)':         Date   Sales   Product
     0 2020-12-01  179.09   Carlota
     1 2020-12-01  161.99   Carlota
     2 2020-12-04  172.46  Sunshine
     3 2020-12-04   64.03      Quad,
     'Data(3)':         Date   Sales  Product
     0 2020-12-02  209.38     Quad
     1 2020-12-04  203.27     Quad
     2 2020-12-04   12.00  Carlota
     3 2020-12-03  112.90  Carlota,
     'Report': Empty DataFrame
     Columns: [Unnamed: 0, Unnamed: 1, Unnamed: 2, Unnamed: 3, Unnamed: 4, Unnamed: 5, Unnamed: 6, Unnamed: 7, Unnamed: 8, Unnamed: 9, Unnamed: 10, Unnamed: 11, Unnamed: 12, FilePath]
     Index: []}




```python
#combine into one dataframe

#list comprehension

combo = [data.assign(sheet=sheetname)
         for sheetname, data
         in all_sheets.items()
        ]

pd.concat(combo,ignore_index=True).dropna(how='all',axis=1)
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
      <th>Date</th>
      <th>Sales</th>
      <th>Product</th>
      <th>sheet</th>
    </tr>
  </thead>
  <tbody>
    <tr>
      <th>0</th>
      <td>2020-12-01</td>
      <td>14.35</td>
      <td>Quad</td>
      <td>Data(1)</td>
    </tr>
    <tr>
      <th>1</th>
      <td>2020-12-02</td>
      <td>144.42</td>
      <td>Sunshine</td>
      <td>Data(1)</td>
    </tr>
    <tr>
      <th>2</th>
      <td>2020-12-04</td>
      <td>207.00</td>
      <td>Quad</td>
      <td>Data(1)</td>
    </tr>
    <tr>
      <th>3</th>
      <td>2020-12-01</td>
      <td>247.33</td>
      <td>Sunshine</td>
      <td>Data(1)</td>
    </tr>
    <tr>
      <th>4</th>
      <td>2020-12-01</td>
      <td>179.09</td>
      <td>Carlota</td>
      <td>Data(2)</td>
    </tr>
    <tr>
      <th>5</th>
      <td>2020-12-01</td>
      <td>161.99</td>
      <td>Carlota</td>
      <td>Data(2)</td>
    </tr>
    <tr>
      <th>6</th>
      <td>2020-12-04</td>
      <td>172.46</td>
      <td>Sunshine</td>
      <td>Data(2)</td>
    </tr>
    <tr>
      <th>7</th>
      <td>2020-12-04</td>
      <td>64.03</td>
      <td>Quad</td>
      <td>Data(2)</td>
    </tr>
    <tr>
      <th>8</th>
      <td>2020-12-02</td>
      <td>209.38</td>
      <td>Quad</td>
      <td>Data(3)</td>
    </tr>
    <tr>
      <th>9</th>
      <td>2020-12-04</td>
      <td>203.27</td>
      <td>Quad</td>
      <td>Data(3)</td>
    </tr>
    <tr>
      <th>10</th>
      <td>2020-12-04</td>
      <td>12.00</td>
      <td>Carlota</td>
      <td>Data(3)</td>
    </tr>
    <tr>
      <th>11</th>
      <td>2020-12-03</td>
      <td>112.90</td>
      <td>Carlota</td>
      <td>Data(3)</td>
    </tr>
  </tbody>
</table>
</div>


